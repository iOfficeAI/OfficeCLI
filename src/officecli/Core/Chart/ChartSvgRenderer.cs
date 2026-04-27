// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// Shared chart SVG rendering logic used by both PowerPoint and Excel HTML preview.
/// Split across two files:
///   ChartSvgRenderer.cs           — regular c:chart extraction + render
///   ChartSvgRenderer.CxExtract.cs — cx:chart extraction + render (histogram,
///                                    funnel, treemap, sunburst, boxWhisker)
/// </summary>
internal partial class ChartSvgRenderer
{
    // CONSISTENCY(chart-default-palette): canonical source is
    // OfficeDefaultThemeColors.DefaultChartSeriesPalette; SVG just needs
    // the '#'-prefixed form, so we derive once at static init.
    public static readonly string[] FallbackColors =
        OfficeDefaultThemeColors.DefaultChartSeriesPalette
            .Select(hex => "#" + hex)
            .ToArray();

    /// <summary>
    /// Theme-derived accent colors for chart series. Set from document theme accent1-6.
    /// Falls back to FallbackColors if not set.
    /// </summary>
    public string[]? ThemeAccentColors { get; set; }

    /// <summary>Get effective default colors: theme accents (with shade/tint variants) or fallback.</summary>
    public string[] DefaultColors => ThemeAccentColors ?? FallbackColors;

    /// <summary>Build theme accent color array from theme color map (accent1-6 + shade variants).</summary>
    public static string[] BuildThemeAccentColors(Dictionary<string, string> themeColors)
    {
        var accents = new List<string>();
        for (int i = 1; i <= 6; i++)
        {
            if (themeColors.TryGetValue($"accent{i}", out var hex))
                accents.Add($"#{hex}");
            else
                accents.Add(FallbackColors[(i - 1) % FallbackColors.Length]);
        }
        // Generate shade variants for cycling (darker versions of accent1-6)
        foreach (var accent in accents.ToList())
        {
            var raw = accent.TrimStart('#');
            accents.Add(ColorMath.ApplyTransforms(raw, shade: 50000)); // 50% shade
        }
        return accents.ToArray();
    }

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
        bool showDataLabels = false, string? valNumFmt = null, string? plotFillColor = null,
        List<(string Name, double Value, string Color, double WidthPt, string Dash)>? referenceLines = null,
        bool isWaterfall = false, List<ErrorBarInfo?>? errorBars = null)
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
                        // Label at segment center — skip if segment narrower than ~2 chars to avoid overflow
                        if (showDataLabels && barW > DataLabelFontPx * 1.6)
                        {
                            var vlabel = rawVal % 1 == 0 ? $"{(int)rawVal}" : $"{rawVal:0.#}";
                            sb.AppendLine($"        <text x=\"{bx + barW / 2:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"{ValueColor}\" font-size=\"{DataLabelFontPx}\" text-anchor=\"middle\" dominant-baseline=\"middle\">{vlabel}</text>");
                        }
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
            // Reference-line overlays: horizontal bars → vertical line at value position on the X (value) axis.
            // For percentStacked charts, the value axis is 0–1 in OOXML but we display 0–100, so scale accordingly.
            if (referenceLines != null)
                foreach (var rl in referenceLines)
                {
                    var v = percentStacked ? rl.Value * 100 : rl.Value;
                    if (v < 0 || v > niceMax) continue;
                    var rx = plotOx + (v / niceMax) * plotPw;
                    var strokeColor = rl.Color.StartsWith("#") ? rl.Color : "#" + rl.Color;
                    var dashArray = RefLineDashArray(rl.Dash);
                    sb.AppendLine($"        <line x1=\"{rx:0.#}\" y1=\"{oy}\" x2=\"{rx:0.#}\" y2=\"{oy + ph}\" stroke=\"{strokeColor}\" stroke-width=\"{rl.WidthPt:0.##}\" stroke-dasharray=\"{dashArray}\"/>");
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

            // Track waterfall connector positions for drawing connecting lines
            var wfPrevTopY = double.NaN;

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
                        // For waterfall: skip rendering Base series (s=0), only render Increase/Decrease
                        if (!isWaterfall || s > 0)
                        {
                            if (barH > 0.5)
                                sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                            if (showDataLabels && barH > DataLabelFontPx + 2)
                            {
                                var vlabel = FormatAxisValue(rawVal, valNumFmt);
                                sb.AppendLine($"        <text x=\"{bx + barW / 2:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"{ValueColor}\" font-size=\"{DataLabelFontPx}\" text-anchor=\"middle\" dominant-baseline=\"middle\">{vlabel}</text>");
                            }
                        }
                        // Waterfall connector line from previous bar's top to this bar's top
                        if (isWaterfall && s == 0 && c > 0 && !double.IsNaN(wfPrevTopY))
                        {
                            var connY = oy + ph - (stackY / niceMax) * ph;
                            var prevBx = ox + (c - 1) * groupW + gap + barW;
                            sb.AppendLine($"        <line x1=\"{prevBx:0.#}\" y1=\"{wfPrevTopY:0.#}\" x2=\"{bx:0.#}\" y2=\"{connY:0.#}\" stroke=\"{GridColor}\" stroke-width=\"1\" stroke-dasharray=\"3,2\"/>");
                        }
                        stackY += val;
                    }
                    else
                    {
                        var bx = ox + c * groupW + gap + s * barW;
                        var by = oy + ph - barH;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        if (showDataLabels)
                        {
                            var vlabel = FormatAxisValue(rawVal, valNumFmt);
                            sb.AppendLine($"        <text x=\"{bx + barW / 2:0.#}\" y=\"{by - 3:0.#}\" fill=\"{ValueColor}\" font-size=\"{DataLabelFontPx}\" text-anchor=\"middle\">{vlabel}</text>");
                        }
                    }
                }
                // Track waterfall top position for connector line
                if (isWaterfall)
                    wfPrevTopY = oy + ph - (stackY / niceMax) * ph;
            }
            // Error bars on vertical (column) bar charts
            if (errorBars != null && !stacked)
            {
                for (int s = 0; s < serCount; s++)
                {
                    var eb = s < errorBars.Count ? errorBars[s] : null;
                    if (eb == null) continue;
                    var ebColor = eb.Color ?? "#333";
                    var capW = Math.Max(2, barW * 0.3);
                    double errAmount = eb.Value;
                    if (eb.ValueType is "stdDev" or "stdErr")
                    {
                        var vals = series[s].values;
                        var mean = vals.Average();
                        var variance = vals.Sum(v => (v - mean) * (v - mean)) / vals.Length;
                        var stddev = Math.Sqrt(variance);
                        errAmount = eb.ValueType == "stdErr" ? stddev / Math.Sqrt(vals.Length) : stddev;
                    }
                    for (int c = 0; c < catCount; c++)
                    {
                        var rawVal = c < series[s].values.Length ? series[s].values[c] : 0;
                        var bx = ox + c * groupW + gap + s * barW + barW / 2;
                        var byTop = oy + ph - (rawVal / niceMax) * ph;
                        double plusErr = eb.ValueType == "percentage" ? Math.Abs(rawVal) * eb.Value / 100.0 : errAmount;
                        double minusErr = plusErr;
                        var showPlus = eb.BarType is "both" or "plus";
                        var showMinus = eb.BarType is "both" or "minus";
                        var yTop = showPlus ? oy + ph - ((rawVal + plusErr) / niceMax) * ph : byTop;
                        var yBot = showMinus ? oy + ph - ((rawVal - minusErr) / niceMax) * ph : byTop;
                        sb.AppendLine($"        <line x1=\"{bx:0.#}\" y1=\"{yTop:0.#}\" x2=\"{bx:0.#}\" y2=\"{yBot:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                        if (showPlus)
                            sb.AppendLine($"        <line x1=\"{bx - capW:0.#}\" y1=\"{yTop:0.#}\" x2=\"{bx + capW:0.#}\" y2=\"{yTop:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                        if (showMinus)
                            sb.AppendLine($"        <line x1=\"{bx - capW:0.#}\" y1=\"{yBot:0.#}\" x2=\"{bx + capW:0.#}\" y2=\"{yBot:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
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
            // Reference-line overlays: vertical bars/columns → horizontal line at value position on the Y (value) axis.
            if (referenceLines != null)
                foreach (var rl in referenceLines)
                {
                    var v = percentStacked ? rl.Value * 100 : rl.Value;
                    if (v < 0 || v > niceMax) continue;
                    var ry = oy + ph - (v / niceMax) * ph;
                    var strokeColor = rl.Color.StartsWith("#") ? rl.Color : "#" + rl.Color;
                    var dashArray = RefLineDashArray(rl.Dash);
                    sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{ry:0.#}\" x2=\"{ox + pw}\" y2=\"{ry:0.#}\" stroke=\"{strokeColor}\" stroke-width=\"{rl.WidthPt:0.##}\" stroke-dasharray=\"{dashArray}\"/>");
                }
        }
    }

    public void RenderLineChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph,
        bool showDataLabels = false, List<string>? markerShapes = null, List<int>? markerSizes = null,
        double? logBase = null, bool isReversed = false,
        bool hasDropLines = false, bool hasHighLowLines = false, bool hasUpDownBars = false,
        string? upBarColor = null, string? downBarColor = null,
        double? axisMin = null, double? axisMax = null, double? majorUnit = null, string? valNumFmt = null,
        List<(string Name, double Value, string Color, double WidthPt, string Dash)>? referenceLines = null,
        List<bool>? smooth = null, List<string>? lineDashes = null, List<double>? lineWidths = null,
        string? dropLineColor = null, double dropLineWidth = 0.7, string? dropLineDash = null,
        string? highLowLineColor = null, double highLowLineWidth = 1,
        List<TrendlineInfo?>? trendlines = null, List<ErrorBarInfo?>? errorBars = null)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var dataMax = allValues.Max();
        var dataMin = allValues.Where(v => v > 0).DefaultIfEmpty(1).Min();
        if (dataMax <= 0) dataMax = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        bool isLog = logBase.HasValue && logBase.Value > 1;

        // Compute axis scale
        double niceMax, niceMin, tickStep;
        int nTicks;
        if (isLog)
        {
            var logB = logBase!.Value;
            niceMin = Math.Floor(Math.Log(dataMin) / Math.Log(logB));
            niceMax = Math.Ceiling(Math.Log(dataMax) / Math.Log(logB));
            if (niceMin >= niceMax) niceMax = niceMin + 1;
            nTicks = (int)(niceMax - niceMin);
            tickStep = 1;
        }
        else
        {
            var computeMax = axisMax ?? dataMax;
            (niceMax, tickStep, nTicks) = ComputeNiceAxis(computeMax);
            if (axisMax.HasValue) niceMax = axisMax.Value;
            niceMin = axisMin ?? 0;
            if (majorUnit.HasValue && majorUnit.Value > 0)
            {
                tickStep = majorUnit.Value;
                nTicks = (int)Math.Ceiling((niceMax - niceMin) / tickStep);
            }
        }

        // Value-to-Y mapping
        double MapY(double val)
        {
            double ratio;
            if (isLog)
            {
                var logB = logBase!.Value;
                var logVal = val > 0 ? Math.Log(val) / Math.Log(logB) : niceMin;
                ratio = (logVal - niceMin) / (niceMax - niceMin);
            }
            else
            {
                ratio = (niceMax - niceMin) > 0 ? (val - niceMin) / (niceMax - niceMin) : 0;
            }
            ratio = Math.Max(0, Math.Min(1, ratio));
            return isReversed ? oy + ratio * ph : oy + ph - ratio * ph;
        }

        // Gridlines
        for (int t = 1; t <= nTicks; t++)
        {
            double tickVal = isLog ? niceMin + t : niceMin + tickStep * t;
            var gy = MapY(isLog ? Math.Pow(logBase!.Value, tickVal) : tickVal);
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\" stroke-dasharray=\"none\"/>");
        }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        // Compute all point coordinates first (needed for high-low/up-down)
        var allPoints = new List<List<(double x, double y, double val)>>();
        for (int s = 0; s < series.Count; s++)
        {
            var pts = new List<(double x, double y, double val)>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = MapY(series[s].values[c]);
                pts.Add((px, py, series[s].values[c]));
            }
            allPoints.Add(pts);
        }

        // High-low lines (vertical line from highest to lowest value at each category)
        if (hasHighLowLines && series.Count >= 2)
        {
            for (int c = 0; c < catCount; c++)
            {
                var yVals = allPoints.Where(p => c < p.Count).Select(p => p[c].y).ToArray();
                if (yVals.Length >= 2)
                {
                    var px = allPoints[0][c].x;
                    var hlColor = highLowLineColor ?? "#666";
                    sb.AppendLine($"        <line x1=\"{px:0.#}\" y1=\"{yVals.Min():0.#}\" x2=\"{px:0.#}\" y2=\"{yVals.Max():0.#}\" stroke=\"{hlColor}\" stroke-width=\"{highLowLineWidth:0.#}\"/>");
                }
            }
        }

        // Up-down bars (between first and last series at each category)
        if (hasUpDownBars && series.Count >= 2)
        {
            var barW = Math.Max(4, pw / catCount * 0.4);
            for (int c = 0; c < catCount; c++)
            {
                if (c >= allPoints[0].Count || c >= allPoints[^1].Count) continue;
                var first = allPoints[0][c];
                var last = allPoints[^1][c];
                var isUp = first.val <= last.val;
                var color = isUp ? (upBarColor ?? "4CAF50") : (downBarColor ?? "F44336");
                if (!color.StartsWith("#")) color = "#" + color;
                var topY = Math.Min(first.y, last.y);
                var botY = Math.Max(first.y, last.y);
                var h = Math.Max(1, botY - topY);
                sb.AppendLine($"        <rect x=\"{first.x - barW / 2:0.#}\" y=\"{topY:0.#}\" width=\"{barW:0.#}\" height=\"{h:0.#}\" fill=\"{color}\" stroke=\"#333\" stroke-width=\"0.5\"/>");
            }
        }

        // Draw lines and markers
        for (int s = 0; s < series.Count; s++)
        {
            var pts = allPoints[s];
            if (pts.Count == 0) continue;
            var lineColor = colors[s % colors.Count];
            var isSmooth = smooth != null && s < smooth.Count && smooth[s];
            var dashName = lineDashes != null && s < lineDashes.Count ? lineDashes[s] : "solid";
            var dashAttr = dashName != "solid" ? $" stroke-dasharray=\"{RefLineDashArray(dashName)}\"" : "";
            var lw = lineWidths != null && s < lineWidths.Count ? lineWidths[s] : 2;

            if (isSmooth && pts.Count >= 2)
            {
                // Catmull-Rom to cubic Bezier smooth path
                var d = new StringBuilder();
                d.Append($"M{pts[0].x:0.#},{pts[0].y:0.#}");
                for (int i = 0; i < pts.Count - 1; i++)
                {
                    var p0 = i > 0 ? pts[i - 1] : pts[i];
                    var p1 = pts[i];
                    var p2 = pts[i + 1];
                    var p3 = i + 2 < pts.Count ? pts[i + 2] : pts[i + 1];
                    var cp1x = p1.x + (p2.x - p0.x) / 6.0;
                    var cp1y = p1.y + (p2.y - p0.y) / 6.0;
                    var cp2x = p2.x - (p3.x - p1.x) / 6.0;
                    var cp2y = p2.y - (p3.y - p1.y) / 6.0;
                    d.Append($" C{cp1x:0.#},{cp1y:0.#} {cp2x:0.#},{cp2y:0.#} {p2.x:0.#},{p2.y:0.#}");
                }
                sb.AppendLine($"        <path d=\"{d}\" fill=\"none\" stroke=\"{lineColor}\" stroke-width=\"{lw:0.#}\"{dashAttr}/>");
            }
            else
            {
                var pointStr = string.Join(" ", pts.Select(p => $"{p.x:0.#},{p.y:0.#}"));
                sb.AppendLine($"        <polyline points=\"{pointStr}\" fill=\"none\" stroke=\"{lineColor}\" stroke-width=\"{lw:0.#}\"{dashAttr}/>");
            }

            // Drop lines (vertical from each data point down to X axis)
            if (hasDropLines)
            {
                var baseY = isReversed ? oy : oy + ph;
                var dlColor = dropLineColor ?? "#888";
                var dlDash = dropLineDash != null ? RefLineDashArray(dropLineDash) : "3,2";
                foreach (var pt in pts)
                    sb.AppendLine($"        <line x1=\"{pt.x:0.#}\" y1=\"{pt.y:0.#}\" x2=\"{pt.x:0.#}\" y2=\"{baseY}\" stroke=\"{dlColor}\" stroke-width=\"{dropLineWidth:0.#}\" stroke-dasharray=\"{dlDash}\"/>");
            }

            var shape = markerShapes != null && s < markerShapes.Count ? markerShapes[s] : "circle";
            var mSize = markerSizes != null && s < markerSizes.Count ? markerSizes[s] * 0.6 : 3;
            for (int p = 0; p < pts.Count; p++)
            {
                sb.AppendLine($"        {RenderMarkerSvg(shape, pts[p].x, pts[p].y, mSize, lineColor)}");
                if (showDataLabels)
                {
                    var val = pts[p].val;
                    var vlabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                    sb.AppendLine($"        <text x=\"{pts[p].x:0.#}\" y=\"{pts[p].y - 6:0.#}\" fill=\"{ValueColor}\" font-size=\"{DataLabelFontPx}\" text-anchor=\"middle\">{vlabel}</text>");
                }
            }
        }

        // Error bars
        if (errorBars != null)
        {
            for (int s = 0; s < series.Count; s++)
            {
                var eb = s < errorBars.Count ? errorBars[s] : null;
                if (eb == null) continue;
                var pts = allPoints[s];
                var ebColor = eb.Color ?? "#666";
                var capW = 4.0; // half-width of the cap line

                // Compute error amount per point
                double errAmount = eb.Value;
                if (eb.ValueType is "stdDev" or "stdErr")
                {
                    var vals = series[s].values;
                    var mean = vals.Average();
                    var variance = vals.Sum(v => (v - mean) * (v - mean)) / vals.Length;
                    var stddev = Math.Sqrt(variance);
                    errAmount = eb.ValueType == "stdErr" ? stddev / Math.Sqrt(vals.Length) : stddev;
                }

                for (int p = 0; p < pts.Count; p++)
                {
                    var val = pts[p].val;
                    double plusErr, minusErr;
                    if (eb.ValueType == "percentage")
                    {
                        plusErr = minusErr = Math.Abs(val) * eb.Value / 100.0;
                    }
                    else
                    {
                        plusErr = minusErr = errAmount;
                    }

                    var showPlus = eb.BarType is "both" or "plus";
                    var showMinus = eb.BarType is "both" or "minus";

                    var yTop = showPlus ? MapY(val + plusErr) : pts[p].y;
                    var yBot = showMinus ? MapY(val - minusErr) : pts[p].y;

                    // Vertical line
                    sb.AppendLine($"        <line x1=\"{pts[p].x:0.#}\" y1=\"{yTop:0.#}\" x2=\"{pts[p].x:0.#}\" y2=\"{yBot:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                    // Top cap
                    if (showPlus)
                        sb.AppendLine($"        <line x1=\"{pts[p].x - capW:0.#}\" y1=\"{yTop:0.#}\" x2=\"{pts[p].x + capW:0.#}\" y2=\"{yTop:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                    // Bottom cap
                    if (showMinus)
                        sb.AppendLine($"        <line x1=\"{pts[p].x - capW:0.#}\" y1=\"{yBot:0.#}\" x2=\"{pts[p].x + capW:0.#}\" y2=\"{yBot:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                }
            }
        }

        // Trendlines
        if (trendlines != null)
        {
            for (int s = 0; s < series.Count; s++)
            {
                var tl = s < trendlines.Count ? trendlines[s] : null;
                if (tl == null) continue;
                var pts = allPoints[s];
                if (pts.Count < 2) continue;
                var lineColor = tl.Color ?? colors[s % colors.Count];
                var dashArr = tl.Dash != "solid" ? $" stroke-dasharray=\"{RefLineDashArray(tl.Dash)}\"" : "";

                // Build x/y data arrays (using category indices as x, values as y)
                var xData = new double[pts.Count];
                var yData = new double[pts.Count];
                for (int i = 0; i < pts.Count; i++)
                {
                    xData[i] = i + 1; // 1-based like Excel
                    yData[i] = series[s].values[i];
                }

                // Compute trendline function
                Func<double, double>? trendFn = null;
                string? eqText = null;
                double rSquared = 0;

                switch (tl.Type)
                {
                    case "linear":
                    {
                        var (slope, intercept) = FitLinear(xData, yData);
                        trendFn = x => slope * x + intercept;
                        eqText = $"y = {slope:0.####}x {(intercept >= 0 ? "+" : "−")} {Math.Abs(intercept):0.####}";
                        rSquared = ComputeRSquared(xData, yData, trendFn);
                        break;
                    }
                    case "exp":
                    {
                        var (a, b) = FitExponential(xData, yData);
                        if (!double.IsNaN(a))
                        {
                            trendFn = x => a * Math.Exp(b * x);
                            eqText = $"y = {a:0.####}e^({b:0.####}x)";
                            rSquared = ComputeRSquared(xData, yData, trendFn);
                        }
                        break;
                    }
                    case "log":
                    {
                        var (a, b) = FitLogarithmic(xData, yData);
                        if (!double.IsNaN(a))
                        {
                            trendFn = x => a * Math.Log(x) + b;
                            eqText = $"y = {a:0.####}ln(x) {(b >= 0 ? "+" : "−")} {Math.Abs(b):0.####}";
                            rSquared = ComputeRSquared(xData, yData, trendFn);
                        }
                        break;
                    }
                    case "poly":
                    {
                        var coeffs = FitPolynomial(xData, yData, tl.Order);
                        if (coeffs != null)
                        {
                            trendFn = x =>
                            {
                                double result = 0;
                                for (int i = 0; i < coeffs.Length; i++)
                                    result += coeffs[i] * Math.Pow(x, i);
                                return result;
                            };
                            var eqParts = new List<string>();
                            for (int i = coeffs.Length - 1; i >= 0; i--)
                            {
                                if (i == 0) eqParts.Add($"{coeffs[i]:0.####}");
                                else if (i == 1) eqParts.Add($"{coeffs[i]:0.####}x");
                                else eqParts.Add($"{coeffs[i]:0.####}x^{i}");
                            }
                            eqText = "y = " + string.Join(" + ", eqParts).Replace("+ -", "− ");
                            rSquared = ComputeRSquared(xData, yData, trendFn);
                        }
                        break;
                    }
                    case "power":
                    {
                        var (a, b) = FitPower(xData, yData);
                        if (!double.IsNaN(a))
                        {
                            trendFn = x => a * Math.Pow(x, b);
                            eqText = $"y = {a:0.####}x^{b:0.####}";
                            rSquared = ComputeRSquared(xData, yData, trendFn);
                        }
                        break;
                    }
                    case "movingAvg":
                    {
                        // Moving average: render as polyline of averaged points
                        var period = Math.Max(2, tl.Period);
                        var maPoints = new List<(double x, double y)>();
                        for (int i = period - 1; i < xData.Length; i++)
                        {
                            double sum = 0;
                            for (int j = 0; j < period; j++) sum += yData[i - j];
                            var avgVal = sum / period;
                            var px = ox + (catCount > 1 ? (double)pw * i / (catCount - 1) : pw / 2.0);
                            var py = MapY(avgVal);
                            maPoints.Add((px, py));
                        }
                        if (maPoints.Count >= 2)
                        {
                            var maPath = string.Join(" ", maPoints.Select(p => $"{p.x:0.#},{p.y:0.#}"));
                            sb.AppendLine($"        <polyline points=\"{maPath}\" fill=\"none\" stroke=\"{lineColor}\" stroke-width=\"{tl.Width:0.#}\"{dashArr}/>");
                        }
                        continue; // no equation/R² for moving average
                    }
                }

                if (trendFn == null) continue;

                // Render trendline curve
                var xMin = xData[0] - tl.Backward;
                var xMax = xData[^1] + tl.Forward;
                var steps = 50;
                var tlPoints = new List<(double px, double py)>();
                for (int i = 0; i <= steps; i++)
                {
                    var x = xMin + (xMax - xMin) * i / steps;
                    var y = trendFn(x);
                    if (double.IsNaN(y) || double.IsInfinity(y)) continue;
                    // Map x to pixel: x is 1-based category index
                    var px = ox + (catCount > 1 ? pw * (x - 1) / (catCount - 1) : pw / 2.0);
                    var py = MapY(y);
                    tlPoints.Add((px, py));
                }

                if (tlPoints.Count >= 2)
                {
                    var pathStr = string.Join(" ", tlPoints.Select(p => $"{p.px:0.#},{p.py:0.#}"));
                    sb.AppendLine($"        <polyline points=\"{pathStr}\" fill=\"none\" stroke=\"{lineColor}\" stroke-width=\"{tl.Width:0.#}\"{dashArr}/>");
                }

                // Equation / R² label
                if (tl.DisplayEquation || tl.DisplayRSquared)
                {
                    var labelParts = new List<string>();
                    if (tl.DisplayEquation && eqText != null) labelParts.Add(eqText);
                    if (tl.DisplayRSquared) labelParts.Add($"R² = {rSquared:0.####}");
                    var label = string.Join("  ", labelParts);
                    // Position label near the end of the trendline
                    var labelX = tlPoints.Count > 0 ? tlPoints[^1].px - 4 : ox + pw;
                    var labelY = tlPoints.Count > 0 ? tlPoints[^1].py - 8 : oy + 12;
                    sb.AppendLine($"        <text x=\"{labelX:0.#}\" y=\"{labelY:0.#}\" fill=\"{lineColor}\" font-size=\"8\" text-anchor=\"end\" font-style=\"italic\">{HtmlEncode(label)}</text>");
                }
            }
        }

        // Reference lines
        if (referenceLines != null)
        {
            foreach (var rl in referenceLines)
            {
                var ry = MapY(rl.Value);
                var dashArr = RefLineDashArray(rl.Dash);
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{ry:0.#}\" x2=\"{ox + pw}\" y2=\"{ry:0.#}\" stroke=\"{rl.Color}\" stroke-width=\"{rl.WidthPt:0.#}\" stroke-dasharray=\"{dashArr}\"/>");
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }

        // Value axis labels
        for (int t = 0; t <= nTicks; t++)
        {
            double tickVal;
            string label;
            if (isLog)
            {
                var exp = niceMin + t;
                tickVal = Math.Pow(logBase!.Value, exp);
                label = FormatAxisValue(tickVal, valNumFmt);
            }
            else
            {
                tickVal = niceMin + tickStep * t;
                label = FormatAxisValue(tickVal, valNumFmt);
            }
            var ty = MapY(tickVal);
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    public void RenderPieChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH, double holeRatio = 0.0, bool showDataLabels = false,
        bool showVal = false, bool showPercent = false)
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
                string label;
                if (showVal && !showPercent)
                    label = pct >= 5 ? $"{values[i]:0.##}" : "";
                else if (showPercent && !showVal)
                    label = pct >= 5 ? $"{pct:0}%" : "";
                else if (showVal && showPercent)
                    label = pct >= 5 ? $"{values[i]:0.##} ({pct:0}%)" : "";
                else
                    label = pct >= 5 ? $"{pct:0}%" : ""; // default to percent for pie
                if (!string.IsNullOrEmpty(label))
                    sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"#fff\" font-size=\"{DataLabelFontPx}\" font-weight=\"bold\" text-anchor=\"middle\" dominant-baseline=\"central\">{label}</text>");
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
        string[] categories, List<string> colors, int svgW, int svgH, int catLabelFontSize = 0,
        string radarStyle = "filled")
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
            sb.AppendLine($"        <polygon points=\"{string.Join(" ", gridPoints)}\" fill=\"none\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
        }
        for (int c = 0; c < catCount; c++)
        {
            var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
            sb.AppendLine($"        <line x1=\"{cx:0.#}\" y1=\"{cy:0.#}\" x2=\"{cx + r * Math.Cos(angle):0.#}\" y2=\"{cy + r * Math.Sin(angle):0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
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
                var isFilled = radarStyle == "filled";
                var fillAttr = isFilled ? $"fill=\"{serColor}\" fill-opacity=\"0.2\"" : "fill=\"none\"";
                sb.AppendLine($"        <polygon points=\"{string.Join(" ", points)}\" {fillAttr} stroke=\"{serColor}\" stroke-width=\"2\"/>");
                // Markers for marker and standard styles (standard gets small dots, marker gets circles)
                var showMarkers = radarStyle != "filled";
                var markerR = radarStyle == "marker" ? 4 : 2;
                if (showMarkers)
                {
                    foreach (var pt in points)
                    {
                        var parts = pt.Split(',');
                        sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"{markerR}\" fill=\"{serColor}\"/>");
                    }
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

        var upColor = "#FFFFFF"; var downColor = "#000000"; // OOXML spec defaults
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
        public bool ShowDataLabelVal { get; set; }
        public bool ShowDataLabelPercent { get; set; }
        public double HoleRatio { get; set; }
        public bool IsStacked { get; set; }
        public bool IsPercent { get; set; }
        public bool IsWaterfall { get; set; }
        public bool Is3D { get; set; }
        public int RotateX { get; set; }
        public int RotateY { get; set; }
        public int Perspective { get; set; }
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
        /// <summary>#7f: OOXML c:legendPos InnerText — "b" (bottom, default),
        /// "t" (top), "r" (right), "l" (left), "tr" (top-right). Rendering
        /// adapts the wrapper layout to each position.</summary>
        public string LegendPos { get; set; } = "b";
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
        /// <summary>Reference-line overlays (horizontal dashed lines at constant values).
        /// Filled by ExtractChartInfo from any ref-line-only LineChart in the plot area.</summary>
        public List<(string Name, double Value, string Color, double WidthPt, string Dash)> ReferenceLines { get; set; } = [];

        // --- Marker shapes per series (circle, diamond, square, triangle, star, x, plus, dash, dot, none) ---
        public List<string> MarkerShapes { get; set; } = [];
        public List<int> MarkerSizes { get; set; } = [];

        // --- Smooth line (cubic spline) per series ---
        public List<bool> Smooth { get; set; } = [];

        // --- Dash pattern per series (solid, dash, dot, dashDot, lgDash, etc.) ---
        public List<string> LineDashes { get; set; } = [];

        // --- Line width per series (in points, from a:ln w="...") ---
        public List<double> LineWidths { get; set; } = [];

        // --- Axis features ---
        public double? LogBase { get; set; }
        public bool IsReversed { get; set; }

        // --- Line elements ---
        public bool HasDropLines { get; set; }
        public string? DropLineColor { get; set; }
        public double DropLineWidth { get; set; } = 0.7;
        public string? DropLineDash { get; set; }
        public bool HasHighLowLines { get; set; }
        public string? HighLowLineColor { get; set; }
        public double HighLowLineWidth { get; set; } = 1;
        public bool HasUpDownBars { get; set; }
        public string? UpBarColor { get; set; }
        public string? DownBarColor { get; set; }

        // --- Data table ---
        public bool HasDataTable { get; set; }

        // --- Radar style (standard, marker, filled) ---
        public string RadarStyle { get; set; } = "filled";

        // --- Trendlines per series ---
        public List<TrendlineInfo?> Trendlines { get; set; } = [];

        // --- Error bars per series ---
        public List<ErrorBarInfo?> ErrorBars { get; set; } = [];
    }

    /// <summary>Trendline metadata extracted from OOXML for SVG rendering.</summary>
    public class TrendlineInfo
    {
        public string Type { get; set; } = "linear"; // linear, exp, log, poly, power, movingAvg
        public int Order { get; set; } = 2; // polynomial order
        public int Period { get; set; } = 2; // moving average period
        public double Forward { get; set; } // forward extrapolation
        public double Backward { get; set; } // backward extrapolation
        public double? Intercept { get; set; }
        public bool DisplayEquation { get; set; }
        public bool DisplayRSquared { get; set; }
        public string? Color { get; set; }
        public double Width { get; set; } = 1.5;
        public string Dash { get; set; } = "dash";
    }

    /// <summary>Error bar metadata extracted from OOXML for SVG rendering.</summary>
    public class ErrorBarInfo
    {
        public string ValueType { get; set; } = "fixedValue"; // fixedValue, percentage, stdDev, stdErr
        public string Direction { get; set; } = "y"; // x, y
        public string BarType { get; set; } = "both"; // both, plus, minus
        public double Value { get; set; } = 1; // the error amount
        public string? Color { get; set; }
        public double Width { get; set; } = 1;
    }

    /// <summary>
    /// Remove reference-line overlay series from a data series list, matching the
    /// OOXML series iteration order. Callers that override <see cref="ChartInfo.Series"/>
    /// with locally-resolved data (e.g. ExcelHandler cell-ref resolution) must re-apply
    /// this filter or the ref-line series will be double-rendered as a bar/line segment.
    /// </summary>
    public static List<(string name, double[] values)> FilterReferenceLineSeries(
        OpenXmlElement? plotArea,
        List<(string name, double[] values)> series)
    {
        if (plotArea is not PlotArea pa || series.Count == 0) return series;
        var mask = ChartHelper.ReadReferenceLineMask(pa);
        if (!mask.Any(m => m)) return series;
        return series.Where((_, i) => i >= mask.Count || !mask[i]).ToList();
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
        info.ReferenceLines = ChartHelper.ReadReferenceLines(info.PlotArea);

        // Filter reference-line series out of the renderer's data series list. They
        // are drawn as overlays via info.ReferenceLines so they must not contribute to
        // axis scale, stacking, colors, or legend. ReadAllSeries itself stays inclusive
        // so the user-facing Get()/Query() path continues to surface ref-line series.
        info.Series = FilterReferenceLineSeries(info.PlotArea, info.Series);

        if (info.Series.Count == 0 && info.ReferenceLines.Count == 0) return info;

        info.Is3D = info.ChartType.Contains("3d");
        info.IsWaterfall = info.ChartType == "waterfall";
        info.IsStacked = info.ChartType.Contains("stacked") || info.ChartType.Contains("Stacked") || info.IsWaterfall;
        info.IsPercent = info.ChartType.Contains("percent") || info.ChartType.Contains("Percent");

        // View3D parameters
        if (chart != null)
        {
            var view3dEl = chart.Elements().FirstOrDefault(e => e.LocalName == "view3D");
            if (view3dEl != null)
            {
                var rotXEl = view3dEl.Elements().FirstOrDefault(e => e.LocalName == "rotX");
                var rotYEl = view3dEl.Elements().FirstOrDefault(e => e.LocalName == "rotY");
                var perspEl = view3dEl.Elements().FirstOrDefault(e => e.LocalName == "perspective");
                if (rotXEl != null && int.TryParse(rotXEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var rx)) info.RotateX = rx;
                if (rotYEl != null && int.TryParse(rotYEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var ry)) info.RotateY = ry;
                if (perspEl != null && int.TryParse(perspEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var pv)) info.Perspective = pv;
            }
        }

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
            bool IsOn(string name) => dLbls.Elements().Any(e =>
                e.LocalName == name && e.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == "1");
            info.ShowDataLabelVal = IsOn("showVal");
            info.ShowDataLabelPercent = IsOn("showPercent");
            info.ShowDataLabels = info.ShowDataLabelVal || info.ShowDataLabelPercent || IsOn("showCatName");
        }

        // Doughnut hole size
        if (info.ChartType.Contains("doughnut"))
        {
            var holeSizeEl = chartTypeEl?.Elements().FirstOrDefault(e => e.LocalName == "holeSize");
            var holeSizeVal = holeSizeEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            info.HoleRatio = (holeSizeVal != null && int.TryParse(holeSizeVal, out var hs) ? hs : 10) / 100.0; // OOXML spec default: 10%
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

            // Log scale
            var logBaseEl = scaling?.Elements().FirstOrDefault(e => e.LocalName == "logBase");
            if (logBaseEl != null && double.TryParse(logBaseEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var lb))
                info.LogBase = lb;

            // Axis orientation (reversed)
            var orientEl = scaling?.Elements().FirstOrDefault(e => e.LocalName == "orientation");
            var orientVal = orientEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            info.IsReversed = orientVal == "maxMin";

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
            var legendRPr = legendEl.Descendants<Drawing.RunProperties>().FirstOrDefault()
                ?? (OpenXmlElement?)legendEl.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
            var legendFontSize = legendRPr?.GetAttributes().FirstOrDefault(a => a.LocalName == "sz").Value;
            if (legendFontSize != null && int.TryParse(legendFontSize, out var lfs))
                info.LegendFontSize = $"{lfs / 100.0:0.##}pt";
            info.LegendFontColor = ExtractFontColor(legendRPr);
            // #7f: honor <c:legendPos w:val="r|l|t|b|tr"/>.
            var posEl = legendEl.Elements().FirstOrDefault(e => e.LocalName == "legendPos");
            var posVal = posEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (!string.IsNullOrEmpty(posVal)) info.LegendPos = posVal!;
        }
        else
        {
            info.HasLegend = info.Series.Count > 1 || isPieType || info.ReferenceLines.Count > 0;
        }

        // Marker shapes, smooth, and dash per series
        if (chartTypeEl != null)
        {
            // Chart-level smooth (lineChart > smooth val="1")
            var chartSmooth = chartTypeEl.Elements().FirstOrDefault(e => e.LocalName == "smooth");
            var chartSmoothVal = chartSmooth?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            var chartIsSmooth = chartSmoothVal == "1" || chartSmoothVal == "true";

            foreach (var ser in serElements)
            {
                var marker = ser.Elements().FirstOrDefault(e => e.LocalName == "marker");
                var symbol = marker?.Elements().FirstOrDefault(e => e.LocalName == "symbol");
                var symbolVal = symbol?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value ?? "circle";
                info.MarkerShapes.Add(symbolVal);
                var sizeEl = marker?.Elements().FirstOrDefault(e => e.LocalName == "size");
                var sizeVal = sizeEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                info.MarkerSizes.Add(sizeVal != null && int.TryParse(sizeVal, out var ms) ? ms : 5);

                // Per-series smooth (overrides chart-level)
                var serSmooth = ser.Elements().FirstOrDefault(e => e.LocalName == "smooth");
                var serSmoothVal = serSmooth?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                info.Smooth.Add(serSmooth != null
                    ? (serSmoothVal == "1" || serSmoothVal == "true")
                    : chartIsSmooth);

                // Per-series dash pattern and line width
                var spPr = ser.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                var ln = spPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
                var prstDash = ln?.Elements().FirstOrDefault(e => e.LocalName == "prstDash");
                var dashVal = prstDash?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                info.LineDashes.Add(dashVal ?? "solid");

                // Per-series line width (a:ln w="..." in EMU, convert to pt: 1pt = 12700 EMU)
                var lnWidth = ln?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value;
                info.LineWidths.Add(lnWidth != null && int.TryParse(lnWidth, out var lw) ? Math.Round(lw / 12700.0, 1) : 2);

                // Per-series trendline
                var trendlineEl = ser.Elements().FirstOrDefault(e => e.LocalName == "trendline");
                if (trendlineEl != null)
                {
                    var tlInfo = new TrendlineInfo();
                    var tlType = trendlineEl.Elements().FirstOrDefault(e => e.LocalName == "trendlineType");
                    tlInfo.Type = tlType?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value ?? "linear";
                    var polyOrder = trendlineEl.Elements().FirstOrDefault(e => e.LocalName == "order");
                    if (polyOrder != null && int.TryParse(polyOrder.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var po))
                        tlInfo.Order = po;
                    var period = trendlineEl.Elements().FirstOrDefault(e => e.LocalName == "period");
                    if (period != null && int.TryParse(period.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var per))
                        tlInfo.Period = per;
                    var fwd = trendlineEl.Elements().FirstOrDefault(e => e.LocalName == "forward");
                    if (fwd != null && double.TryParse(fwd.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value,
                        System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var fv))
                        tlInfo.Forward = fv;
                    var bwd = trendlineEl.Elements().FirstOrDefault(e => e.LocalName == "backward");
                    if (bwd != null && double.TryParse(bwd.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value,
                        System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var bv))
                        tlInfo.Backward = bv;
                    var intercept = trendlineEl.Elements().FirstOrDefault(e => e.LocalName == "intercept");
                    if (intercept != null && double.TryParse(intercept.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value,
                        System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var iv))
                        tlInfo.Intercept = iv;
                    var dispEq = trendlineEl.Elements().FirstOrDefault(e => e.LocalName == "dispEq");
                    tlInfo.DisplayEquation = dispEq?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == "1";
                    var dispRSqr = trendlineEl.Elements().FirstOrDefault(e => e.LocalName == "dispRSqr");
                    tlInfo.DisplayRSquared = dispRSqr?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == "1";
                    // Trendline styling
                    var tlSpPr = trendlineEl.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                    var tlLn = tlSpPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
                    tlInfo.Color = ExtractLineColor(tlSpPr);
                    if (tlLn?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value is string tlw
                        && int.TryParse(tlw, out var tlwPt))
                        tlInfo.Width = Math.Round(tlwPt / 12700.0, 1);
                    var tlDash = tlLn?.Elements().FirstOrDefault(e => e.LocalName == "prstDash");
                    tlInfo.Dash = tlDash?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value ?? "dash";
                    info.Trendlines.Add(tlInfo);
                }
                else
                    info.Trendlines.Add(null);

                // Per-series error bars
                var errBarsEl = ser.Elements().FirstOrDefault(e => e.LocalName == "errBars");
                if (errBarsEl != null)
                {
                    var ebInfo = new ErrorBarInfo();
                    var ebType = errBarsEl.Elements().FirstOrDefault(e => e.LocalName == "errValType");
                    ebInfo.ValueType = ebType?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value ?? "fixedValue";
                    var ebDir = errBarsEl.Elements().FirstOrDefault(e => e.LocalName == "errDir");
                    ebInfo.Direction = ebDir?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value ?? "y";
                    var ebBarType = errBarsEl.Elements().FirstOrDefault(e => e.LocalName == "errBarType");
                    ebInfo.BarType = ebBarType?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value ?? "both";
                    // Read error value from Plus/Minus > NumLit > NumericPoint > v
                    var plusEl = errBarsEl.Elements().FirstOrDefault(e => e.LocalName == "plus");
                    var numPt = plusEl?.Descendants().FirstOrDefault(e => e.LocalName == "v");
                    if (numPt != null && double.TryParse(numPt.InnerText,
                        System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var ebVal))
                        ebInfo.Value = ebVal;
                    // Error bar styling
                    var ebSpPr = errBarsEl.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                    ebInfo.Color = ExtractLineColor(ebSpPr);
                    var ebLn = ebSpPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
                    if (ebLn?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value is string ebw
                        && int.TryParse(ebw, out var ebwPt))
                        ebInfo.Width = Math.Round(ebwPt / 12700.0, 1);
                    info.ErrorBars.Add(ebInfo);
                }
                else
                    info.ErrorBars.Add(null);
            }

            // Line elements: dropLines, hiLowLines, upDownBars
            var dropLinesEl = chartTypeEl.Elements().FirstOrDefault(e => e.LocalName == "dropLines");
            info.HasDropLines = dropLinesEl != null;
            if (dropLinesEl != null)
            {
                var dlSpPr = dropLinesEl.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                var dlLn = dlSpPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
                info.DropLineColor = ExtractLineColor(dlSpPr);
                if (dlLn?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value is string dlw
                    && int.TryParse(dlw, out var dlwPt))
                    info.DropLineWidth = Math.Round(dlwPt / 12700.0, 1);
                var dlDash = dlLn?.Elements().FirstOrDefault(e => e.LocalName == "prstDash");
                info.DropLineDash = dlDash?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            }
            var hiLowEl = chartTypeEl.Elements().FirstOrDefault(e => e.LocalName == "hiLowLines");
            info.HasHighLowLines = hiLowEl != null;
            if (hiLowEl != null)
            {
                var hlSpPr = hiLowEl.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                var hlLn = hlSpPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
                info.HighLowLineColor = ExtractLineColor(hlSpPr);
                if (hlLn?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value is string hlw
                    && int.TryParse(hlw, out var hlwPt))
                    info.HighLowLineWidth = Math.Round(hlwPt / 12700.0, 1);
            }
            var upDownBars = chartTypeEl.Elements().FirstOrDefault(e => e.LocalName == "upDownBars");
            info.HasUpDownBars = upDownBars != null;
            if (upDownBars != null)
            {
                var upSpPr = upDownBars.Elements().FirstOrDefault(e => e.LocalName == "upBars")
                    ?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                var dnSpPr = upDownBars.Elements().FirstOrDefault(e => e.LocalName == "downBars")
                    ?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                info.UpBarColor = ExtractFillColor(upSpPr) ?? "4CAF50";
                info.DownBarColor = ExtractFillColor(dnSpPr) ?? "F44336";
            }
        }

        // Data table
        var dataTableEl = chart?.Descendants().FirstOrDefault(e => e.LocalName == "dTable");
        info.HasDataTable = dataTableEl != null;

        // Radar style
        var radarChartEl = plotArea.Elements().FirstOrDefault(e => e.LocalName == "radarChart");
        if (radarChartEl != null)
        {
            var rsEl = radarChartEl.Elements().FirstOrDefault(e => e.LocalName == "radarStyle");
            var rsVal = rsEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            info.RadarStyle = rsVal ?? "marker";
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
                colors.Add(rgb != null ? $"#{rgb}" : FallbackColors[i % FallbackColors.Length]);
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
                colors.Add(rgb != null ? $"#{rgb}" : FallbackColors[i % FallbackColors.Length]);
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
        var v = srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        // Reject non-hex values — the return flows into $"#{...}" inline SVG
        // fill/style attributes. Same XSS class as w:color / w:shd / border.
        if (v == null) return null;
        if (v.Length is not (3 or 6 or 8)) return null;
        foreach (var c in v)
            if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'))) return null;
        return v;
    }

    /// <summary>Extract font color from RunProperties or DefaultRunProperties (solidFill > srgbClr).</summary>
    private static string? ExtractFontColor(OpenXmlElement? rPr)
    {
        if (rPr == null) return null;
        var solidFill = rPr.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        var srgb = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        var val = srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        return HexOrNull(val);
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
        return HexOrNull(val);
    }

    // Hex-only stripper: reject non-hex so these chart-color getters can't
    // become XSS sinks when their return flows into SVG style/fill/stroke
    // attributes downstream in Excel/PPTX/Word previews.
    private static string? HexOrNull(string? v)
    {
        if (v == null) return null;
        if (v.Length is not (3 or 6 or 8)) return null;
        foreach (var c in v)
            if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'))) return null;
        return v;
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

        // cx extended chart types (funnel / treemap / sunburst / boxWhisker)
        // dispatch to dedicated emitters before the regular bar/line/pie
        // branches — otherwise they fall through to the column fallback and
        // render as generic bar charts. Histogram intentionally falls through
        // here: it uses the regular column pipeline after ExtractCxChartInfo
        // has pre-binned the values into categories.
        if (TryRenderCxSpecificType(sb, info, marginLeft, marginTop, plotW, plotH))
            return;

        if (chartType.Contains("pie") || chartType.Contains("doughnut"))
        {
            if (info.Is3D)
                RenderPie3DSvg(sb, info.Series, info.Categories, info.Colors, svgW, svgH,
                    info.ShowDataLabels, info.ShowDataLabelVal, info.ShowDataLabelPercent,
                    info.RotateX > 0 ? info.RotateX : 30);
            else
                RenderPieChartSvg(sb, info.Series, info.Categories, info.Colors, svgW, svgH, info.HoleRatio, info.ShowDataLabels,
                    info.ShowDataLabelVal, info.ShowDataLabelPercent);
        }
        else if (chartType.Contains("area"))
        {
            var areaW = plotW - (int)(plotW * 0.03);
            if (info.Is3D)
                RenderArea3DSvg(sb, info.Series, info.Categories, info.Colors, marginLeft, marginTop, areaW, plotH,
                    info.IsStacked, info.RotateX, info.RotateY);
            else
                RenderAreaChartSvg(sb, info.Series, info.Categories, info.Colors, marginLeft, marginTop, areaW, plotH, info.IsStacked);
        }
        else if (chartType == "combo")
        {
            RenderComboChartSvg(sb, info.PlotArea!, info.Series, info.Categories, info.Colors, marginLeft, marginTop, plotW, plotH);
        }
        else if (chartType.Contains("radar"))
        {
            RenderRadarChartSvg(sb, info.Series, info.Categories, info.Colors, svgW, svgH, CatFontPx, info.RadarStyle);
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
                RenderLineChartSvg(sb, info.Series, info.Categories, info.Colors, marginLeft, marginTop, plotW, plotH,
                    info.ShowDataLabels, info.MarkerShapes, info.MarkerSizes, info.LogBase, info.IsReversed,
                    info.HasDropLines, info.HasHighLowLines, info.HasUpDownBars,
                    info.UpBarColor, info.DownBarColor, info.AxisMin, info.AxisMax, info.MajorUnit, info.ValNumFmt,
                    info.ReferenceLines, info.Smooth, info.LineDashes, info.LineWidths,
                    info.DropLineColor, info.DropLineWidth, info.DropLineDash,
                    info.HighLowLineColor, info.HighLowLineWidth,
                    info.Trendlines, info.ErrorBars);
        }
        else
        {
            // Column/bar variants
            var isHorizontal = chartType.Contains("bar") && !chartType.Contains("column");
            // Horizontal bars have their own hLabelMargin inside, so reduce outer marginLeft
            var barMarginLeft = isHorizontal ? 5 : marginLeft;
            var barPlotW = isHorizontal ? svgW - barMarginLeft - marginRight : plotW;
            if (info.Is3D)
                RenderBar3DSvg(sb, info.Series, info.Categories, info.Colors, barMarginLeft, marginTop, barPlotW, plotH, isHorizontal,
                    info.IsStacked, info.IsPercent, info.AxisMax, info.AxisMin, info.MajorUnit,
                    info.GapWidth, info.ShowDataLabels, info.ValNumFmt,
                    info.ReferenceLines, info.RotateX, info.RotateY);
            else
                RenderBarChartSvg(sb, info.Series, info.Categories, info.Colors, barMarginLeft, marginTop, barPlotW, plotH,
                    isHorizontal, info.IsStacked, info.IsPercent, info.AxisMax, info.AxisMin, info.MajorUnit,
                    info.GapWidth, ValFontPx, CatFontPx, info.ShowDataLabels, info.ValNumFmt,
                    isHorizontal ? info.PlotFillColor : null, info.ReferenceLines,
                    info.IsWaterfall, info.ErrorBars);
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
        var legendColor = info.LegendFontColor ?? fontColor;
        var isPieType = info.ChartType.Contains("pie") || info.ChartType.Contains("doughnut");
        // #7f: legendPos "r" / "l" / "tr" stack swatches vertically; "b" / "t"
        // keep the horizontal row layout but the caller wraps with flex so
        // they appear above / below the SVG.
        var isVertical = info.LegendPos is "r" or "l" or "tr";
        var layoutCss = isVertical
            ? "display:flex;flex-direction:column;gap:6px;padding:4px 6px;align-items:flex-start"
            : "display:flex;flex-wrap:wrap;justify-content:center;gap:16px;padding:4px 0";
        // Whitelist legendPos: ST_LegendPos values are short tokens, so
        // reject anything outside the schema to stop an adversarial
        // <c:legendPos val='x" onclick=..."'/> from escaping the attr.
        var safePos = info.LegendPos is "r" or "l" or "t" or "b" or "tr" or "ctr" ? info.LegendPos : "";
        sb.Append($"<div class=\"chart-legend\" data-legend-pos=\"{safePos}\" style=\"{layoutCss};font-size:{info.LegendFontSize};color:{legendColor}\">");
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
            // Office convention: horizontal bar charts render legend in reverse of
            // declaration order so stacking reads top-to-bottom matching legend order.
            // CONSISTENCY(chart-legend-order): vertical bar/column, line, area keep
            // declaration order.
            var isHorizBarLegend = info.ChartType.Contains("bar") && !info.ChartType.Contains("column");
            for (int k = 0; k < info.Series.Count; k++)
            {
                int i = isHorizBarLegend ? info.Series.Count - 1 - k : k;
                var color = i < info.Colors.Count ? info.Colors[i] : DefaultColors[i % DefaultColors.Length];
                sb.Append($"<span style=\"display:inline-flex;align-items:center;gap:4px\"><span style=\"display:inline-block;width:12px;height:12px;background:{color};border-radius:1px\"></span>{HtmlEncode(info.Series[i].name)}</span>");
            }
            // Reference-line entries render as a dashed swatch beside the regular series.
            foreach (var rl in info.ReferenceLines)
            {
                var color = rl.Color.StartsWith("#") ? rl.Color : "#" + rl.Color;
                var name = string.IsNullOrEmpty(rl.Name) ? "Ref" : rl.Name;
                sb.Append($"<span style=\"display:inline-flex;align-items:center;gap:4px\"><svg width=\"16\" height=\"10\" style=\"vertical-align:middle\"><line x1=\"0\" y1=\"5\" x2=\"16\" y2=\"5\" stroke=\"{color}\" stroke-width=\"{rl.WidthPt:0.##}\" stroke-dasharray=\"{RefLineDashArray(rl.Dash)}\"/></svg>{HtmlEncode(name)}</span>");
            }
        }
        sb.AppendLine("</div>");
    }

    /// <summary>Render a data table below the chart (HTML table showing raw series values).</summary>
    public void RenderDataTableHtml(StringBuilder sb, ChartInfo info)
    {
        if (!info.HasDataTable) return;
        sb.AppendLine("  <div style=\"overflow-x:auto;padding:0 4px\">");
        sb.AppendLine("  <table style=\"width:100%;border-collapse:collapse;font-size:7pt;color:#555;margin-top:2px\">");
        // Header row: categories
        sb.Append("    <tr><td style=\"border:1px solid #ccc;padding:1px 3px\"></td>");
        foreach (var cat in info.Categories)
            sb.Append($"<td style=\"border:1px solid #ccc;padding:1px 3px;text-align:center;font-weight:bold\">{HtmlEncode(cat)}</td>");
        sb.AppendLine("</tr>");
        // Series rows
        for (int s = 0; s < info.Series.Count; s++)
        {
            var color = s < info.Colors.Count ? info.Colors[s] : DefaultColors[s % DefaultColors.Length];
            sb.Append($"    <tr><td style=\"border:1px solid #ccc;padding:1px 3px;font-weight:bold;color:{color}\">{HtmlEncode(info.Series[s].name)}</td>");
            for (int c = 0; c < info.Categories.Length; c++)
            {
                var val = c < info.Series[s].values.Length ? info.Series[s].values[c] : 0;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                sb.Append($"<td style=\"border:1px solid #ccc;padding:1px 3px;text-align:center\">{label}</td>");
            }
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("  </table>");
        sb.AppendLine("  </div>");
    }

    // ==================== Reference Line Helpers ====================

    /// <summary>Map an OOXML PresetLineDashValues InnerText (e.g. "sysDash", "lgDashDot") to
    /// an SVG stroke-dasharray value. Falls back to a generic dashed pattern for unknowns.</summary>
    private static string RefLineDashArray(string dashName) => dashName.ToLowerInvariant() switch
    {
        "solid" => "none",
        "dot" or "sysdot" => "1,2",
        "dash" or "sysdash" => "5,3",
        "dashdot" or "sysdashdot" => "5,3,1,3",
        "lgdash" or "longdash" => "8,3",
        "lgdashdot" or "longdashdot" => "8,3,1,3",
        "lgdashdotdot" or "longdashdotdot" => "8,3,1,3,1,3",
        _ => "5,3"
    };

    // ==================== 3D Chart Helpers ====================

    /// <summary>Darken or lighten a hex color by a factor (0.0-2.0, 1.0=unchanged)</summary>
    private static string RenderMarkerSvg(string shape, double cx, double cy, double r, string color)
    {
        return shape switch
        {
            "diamond" => $"<polygon points=\"{cx},{cy - r} {cx + r},{cy} {cx},{cy + r} {cx - r},{cy}\" fill=\"{color}\"/>",
            "square" => $"<rect x=\"{cx - r}\" y=\"{cy - r}\" width=\"{r * 2}\" height=\"{r * 2}\" fill=\"{color}\"/>",
            "triangle" => $"<polygon points=\"{cx},{cy - r} {cx + r},{cy + r} {cx - r},{cy + r}\" fill=\"{color}\"/>",
            "star" => BuildStarPath(cx, cy, r, color),
            "x" => $"<g stroke=\"{color}\" stroke-width=\"1.5\"><line x1=\"{cx - r}\" y1=\"{cy - r}\" x2=\"{cx + r}\" y2=\"{cy + r}\"/><line x1=\"{cx + r}\" y1=\"{cy - r}\" x2=\"{cx - r}\" y2=\"{cy + r}\"/></g>",
            "plus" => $"<g stroke=\"{color}\" stroke-width=\"1.5\"><line x1=\"{cx}\" y1=\"{cy - r}\" x2=\"{cx}\" y2=\"{cy + r}\"/><line x1=\"{cx - r}\" y1=\"{cy}\" x2=\"{cx + r}\" y2=\"{cy}\"/></g>",
            "dash" => $"<line x1=\"{cx - r}\" y1=\"{cy}\" x2=\"{cx + r}\" y2=\"{cy}\" stroke=\"{color}\" stroke-width=\"2\"/>",
            "dot" => $"<circle cx=\"{cx}\" cy=\"{cy}\" r=\"1.5\" fill=\"{color}\"/>",
            "none" => "",
            _ => $"<circle cx=\"{cx}\" cy=\"{cy}\" r=\"{r}\" fill=\"{color}\"/>", // circle or auto
        };
    }

    private static string BuildStarPath(double cx, double cy, double r, string color)
    {
        var sb = new StringBuilder();
        sb.Append($"<polygon points=\"");
        for (int i = 0; i < 10; i++)
        {
            var angle = Math.PI / 2 + i * Math.PI / 5;
            var rad = i % 2 == 0 ? r : r * 0.4;
            sb.Append($"{cx + rad * Math.Cos(angle):0.#},{cy - rad * Math.Sin(angle):0.#} ");
        }
        sb.Append($"\" fill=\"{color}\"/>");
        return sb.ToString();
    }

    private static string AdjustColor(string hexColor, double factor)
    {
        var hex = hexColor.TrimStart('#');
        if (hex.Length < 6) return hexColor;
        var r = (int)Math.Clamp(int.Parse(hex[..2], System.Globalization.NumberStyles.HexNumber) * factor, 0, 255);
        var g = (int)Math.Clamp(int.Parse(hex[2..4], System.Globalization.NumberStyles.HexNumber) * factor, 0, 255);
        var b = (int)Math.Clamp(int.Parse(hex[4..6], System.Globalization.NumberStyles.HexNumber) * factor, 0, 255);
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    // 3D isometric offsets (defaults for 0/0 view3D)
    private const double Depth3D = 12;
    private const double DxIso = 8;
    private const double DyIso = -6;

    /// <summary>Compute 3D isometric offsets from view3D parameters.</summary>
    private static (double dx, double dy) Compute3DOffsets(int rotateX, int rotateY, double baseDepth = 10)
    {
        if (rotateX == 0 && rotateY == 0) return (DxIso, DyIso);
        var ry = Math.Clamp(rotateY, 0, 360) * Math.PI / 180;
        var rx = Math.Clamp(rotateX, 0, 90) * Math.PI / 180;
        var dx = baseDepth * Math.Sin(ry) * 0.9;
        var dy = -baseDepth * Math.Sin(rx) * 0.7;
        if (Math.Abs(dx) < 2) dx = dx >= 0 ? 2 : -2;
        if (Math.Abs(dy) < 2) dy = -2;
        return (dx, dy);
    }

    private void RenderBar3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool horizontal,
        bool stacked = false, bool percentStacked = false,
        double? ooxmlMax = null, double? ooxmlMin = null, double? ooxmlMajorUnit = null,
        int? ooxmlGapWidth = null, bool showDataLabels = false, string? valNumFmt = null,
        List<(string Name, double Value, string Color, double WidthPt, string Dash)>? referenceLines = null,
        int rotateX = 15, int rotateY = 20)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var serCount = series.Count;
        var (dx3d, dy3d) = Compute3DOffsets(rotateX, rotateY);

        // Compute axis range (mirrors 2D RenderBarChartSvg logic)
        double maxVal, minVal = 0;
        if (stacked || percentStacked)
        {
            var catSums = new double[catCount];
            for (int c = 0; c < catCount; c++)
                catSums[c] = series.Sum(s => c < s.values.Length ? s.values[c] : 0);
            maxVal = percentStacked ? 100 : catSums.Max();
        }
        else
            maxVal = allValues.Max();

        if (ooxmlMax.HasValue) maxVal = ooxmlMax.Value;
        if (ooxmlMin.HasValue) minVal = ooxmlMin.Value;
        if (maxVal <= minVal) maxVal = minVal + 1;
        var range = maxVal - minVal;

        // Grid ticks
        int tickCount;
        double majorUnit;
        if (ooxmlMajorUnit.HasValue && ooxmlMajorUnit.Value > 0) { majorUnit = ooxmlMajorUnit.Value; tickCount = (int)(range / majorUnit); }
        else { var (nm, _, nu) = ComputeNiceAxis(maxVal); maxVal = nm; range = maxVal - minVal; majorUnit = nu > 0 ? nu : range / 4; tickCount = majorUnit > 0 ? (int)(range / majorUnit) : 4; }

        void Draw3DBar(double bx, double by, double barW2, double barH2, string color)
        {
            if (barH2 < 0.5) return;
            var sideColor = AdjustColor(color, 0.65);
            var topColor = AdjustColor(color, 1.25);
            // Front face
            sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW2:0.#}\" height=\"{barH2:0.#}\" fill=\"{color}\" opacity=\"0.9\"/>");
            // Top face
            sb.AppendLine($"        <polygon points=\"{bx:0.#},{by:0.#} {bx + barW2:0.#},{by:0.#} {bx + barW2 + dx3d:0.#},{by + dy3d:0.#} {bx + dx3d:0.#},{by + dy3d:0.#}\" fill=\"{topColor}\" opacity=\"0.9\"/>");
            // Right side face
            sb.AppendLine($"        <polygon points=\"{bx + barW2:0.#},{by:0.#} {bx + barW2 + dx3d:0.#},{by + dy3d:0.#} {bx + barW2 + dx3d:0.#},{by + barH2 + dy3d:0.#} {bx + barW2:0.#},{by + barH2:0.#}\" fill=\"{sideColor}\" opacity=\"0.9\"/>");
        }

        if (horizontal)
        {
            var maxLabelLen = categories.Length > 0 ? categories.Max(c => c.Length) : 0;
            var hLabelMargin = (int)(maxLabelLen * CatFontPx * 0.5) + 4;
            var plotOx = ox + hLabelMargin;
            var plotPw = pw - hLabelMargin;
            var groupH = (double)ph / Math.Max(catCount, 1);
            var barH = stacked || percentStacked ? groupH * 0.5 : groupH * 0.5 / serCount;
            var gap = groupH * 0.2;

            // Gridlines
            for (int t = 1; t <= tickCount; t++)
            {
                var gx = plotOx + (double)plotPw * t / tickCount;
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
            }
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy}\" x2=\"{plotOx}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy + ph}\" x2=\"{plotOx + plotPw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

            for (int c = 0; c < catCount; c++)
            {
                if (stacked || percentStacked)
                {
                    var catTotal = series.Sum(s => c < s.values.Length ? s.values[c] : 0);
                    double cumX = 0;
                    for (int s = 0; s < serCount; s++)
                    {
                        var val = c < series[s].values.Length ? series[s].values[c] : 0;
                        var normVal = percentStacked && catTotal > 0 ? val / catTotal * 100 : val;
                        var segW = (normVal / range) * plotPw;
                        var by = oy + c * groupH + gap;
                        var color = colors[s % colors.Count];
                        Draw3DBar(plotOx + cumX, by, segW, barH, color);
                        cumX += segW;
                    }
                }
                else
                {
                    for (int s = 0; s < serCount; s++)
                    {
                        if (c >= series[s].values.Length) continue;
                        var val = series[s].values[c];
                        var barW2 = ((val - minVal) / range) * plotPw;
                        var by = oy + c * groupH + gap + s * barH;
                        Draw3DBar(plotOx, by, barW2, barH, colors[s % colors.Count]);
                    }
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                sb.AppendLine($"        <text x=\"{plotOx - 4}\" y=\"{oy + c * groupH + groupH / 2:0.#}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= tickCount; t++)
            {
                var val = minVal + majorUnit * t;
                var label = FormatAxisValue(val, valNumFmt);
                sb.AppendLine($"        <text x=\"{plotOx + (double)plotPw * t / tickCount:0.#}\" y=\"{oy + ph + 16}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"middle\">{label}</text>");
            }
        }
        else
        {
            var gapPct = ooxmlGapWidth.HasValue ? ooxmlGapWidth.Value / 100.0 : 1.5;
            var groupW = (double)pw / Math.Max(catCount, 1);
            double barW;
            if (stacked || percentStacked)
                barW = groupW / (1 + gapPct);
            else
                barW = groupW / (serCount + gapPct);
            var gapW = (groupW - (stacked || percentStacked ? barW : barW * serCount)) / 2;

            // Gridlines
            for (int t = 1; t <= tickCount; t++)
            {
                var gy = oy + ph - (double)ph * t / tickCount;
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
            }
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

            // Reference lines
            if (referenceLines != null)
            {
                foreach (var rl in referenceLines)
                {
                    var rly = oy + ph - ((rl.Value - minVal) / range) * ph;
                    var rlDash = rl.Dash == "dash" ? "stroke-dasharray=\"6,3\"" : rl.Dash == "dot" ? "stroke-dasharray=\"2,2\"" : "";
                    sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{rly:0.#}\" x2=\"{ox + pw}\" y2=\"{rly:0.#}\" stroke=\"{rl.Color}\" stroke-width=\"{rl.WidthPt:0.#}\" {rlDash}/>");
                }
            }

            for (int c = 0; c < catCount; c++)
            {
                if (stacked || percentStacked)
                {
                    var catTotal = series.Sum(s => c < s.values.Length ? s.values[c] : 0);
                    double cumH = 0;
                    for (int s = 0; s < serCount; s++)
                    {
                        var val = c < series[s].values.Length ? series[s].values[c] : 0;
                        var normVal = percentStacked && catTotal > 0 ? val / catTotal * 100 : val;
                        var segH = ((normVal) / range) * ph;
                        var bx = ox + c * groupW + gapW;
                        var by = oy + ph - cumH - segH;
                        Draw3DBar(bx, by, barW, segH, colors[s % colors.Count]);
                        if (showDataLabels && segH > 10)
                        {
                            var vlabel = FormatAxisValue(val, valNumFmt);
                            sb.AppendLine($"        <text x=\"{bx + barW / 2:0.#}\" y=\"{by + segH / 2:0.#}\" fill=\"white\" font-size=\"{DataLabelFontPx}\" text-anchor=\"middle\" dominant-baseline=\"middle\">{vlabel}</text>");
                        }
                        cumH += segH;
                    }
                }
                else
                {
                    for (int s = 0; s < serCount; s++)
                    {
                        if (c >= series[s].values.Length) continue;
                        var val = series[s].values[c];
                        var barH2 = ((val - minVal) / range) * ph;
                        var bx = ox + c * groupW + gapW + s * barW;
                        var by = oy + ph - barH2;
                        Draw3DBar(bx, by, barW, barH2, colors[s % colors.Count]);
                        if (showDataLabels)
                        {
                            var vlabel = FormatAxisValue(val, valNumFmt);
                            sb.AppendLine($"        <text x=\"{bx + barW / 2 + dx3d / 2:0.#}\" y=\"{by + dy3d - 3:0.#}\" fill=\"{ValueColor}\" font-size=\"{DataLabelFontPx}\" text-anchor=\"middle\">{vlabel}</text>");
                        }
                    }
                }
            }
            // Category labels
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                sb.AppendLine($"        <text x=\"{ox + c * groupW + groupW / 2:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
            }
            // Value axis labels
            for (int t = 0; t <= tickCount; t++)
            {
                var val = minVal + majorUnit * t;
                var label = FormatAxisValue(val, valNumFmt);
                var ty = oy + ph - ((val - minVal) / range) * ph;
                sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
            }
        }
    }

    private void RenderPie3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH,
        bool showDataLabels = false, bool showVal = false, bool showPercent = false,
        int rotateX = 30)
    {
        var values = series.FirstOrDefault().values ?? [];
        if (values.Length == 0) return;
        var total = values.Sum();
        if (total <= 0) return;

        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var rx = Math.Min(svgW, svgH) * 0.35;
        // Use rotateX to control squash: higher angle = more tilted = more elliptical
        var tilt = Math.Clamp(rotateX > 0 ? rotateX : 30, 5, 80) * Math.PI / 180;
        var ry = rx * Math.Cos(tilt);
        var depth = rx * 0.08 + rx * 0.12 * (Math.Sin(tilt));
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

        // Side walls — sort by midpoint closeness to PI (front) for correct z-order
        var wallSlices = slices.Where(s => s.start < Math.PI && s.end > 0).OrderBy(s =>
        {
            var mid = (s.start + s.end) / 2;
            return -Math.Abs(mid - Math.PI / 2); // draw furthest from front first
        }).ToList();

        foreach (var (idx, start, end, color) in wallSlices)
        {
            var sideColor = AdjustColor(color, 0.6);
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

        // Top face slices
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

            // Data labels
            var midAngle = startAngle + sliceAngle / 2;
            var labelR = rx * 0.65;
            var lx = cx + labelR * Math.Cos(midAngle);
            var ly = cy + (labelR * Math.Cos(tilt)) * Math.Sin(midAngle);
            var pct = total > 0 ? values[i] / total * 100 : 0;

            if (showDataLabels || showVal || showPercent)
            {
                var parts = new List<string>();
                if (showVal) parts.Add(values[i] % 1 == 0 ? $"{(int)values[i]}" : $"{values[i]:0.#}");
                if (showPercent) parts.Add($"{pct:0}%");
                if (parts.Count == 0) parts.Add($"{pct:0}%"); // default to percent
                var labelText = string.Join("\n", parts);
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"white\" font-size=\"9\" font-weight=\"bold\" text-anchor=\"middle\" dominant-baseline=\"middle\">{HtmlEncode(labelText)}</text>");
            }
            else
            {
                // Category name label
                var catLabel = i < categories.Length ? categories[i] : "";
                if (!string.IsNullOrEmpty(catLabel))
                    sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"white\" font-size=\"9\" text-anchor=\"middle\" dominant-baseline=\"middle\">{HtmlEncode(catLabel)}</text>");
            }

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

    private void RenderArea3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph,
        bool stacked = false, int rotateX = 15, int rotateY = 20)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var serCount = series.Count;

        double maxVal;
        if (stacked)
        {
            var catSums = new double[catCount];
            for (int c = 0; c < catCount; c++)
                catSums[c] = series.Sum(s => c < s.values.Length ? s.values[c] : 0);
            maxVal = catSums.Max();
        }
        else
            maxVal = allValues.Max();
        var (niceMax, _, _) = ComputeNiceAxis(maxVal);
        maxVal = niceMax;
        if (maxVal <= 0) maxVal = 1;

        // 3D layout: reserve space for depth lanes
        // Each series gets a "lane" along the depth (diagonal) direction
        var laneCount = stacked ? 1 : serCount;
        var laneStep = Math.Min(pw, ph) * 0.10; // step between lane starts (includes gap)
        var laneThickness = laneStep * 0.55;     // actual wall thickness (rest is gap)
        var totalDepthX = laneStep * laneCount * 0.7;  // total horizontal depth shift
        var totalDepthY = -laneStep * laneCount * 0.5;  // total vertical depth shift (upward)

        // Shrink front plot area to make room for depth
        var plotW = (int)(pw - totalDepthX);
        var plotH = (int)(ph + totalDepthY); // totalDepthY is negative

        // Axes & gridlines on the front plane
        for (int t = 1; t <= 4; t++)
        {
            var gy = oy + plotH - (double)plotH * t / 4;
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + plotW}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
        }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + totalDepthY}\" x2=\"{ox}\" y2=\"{oy + plotH}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + plotH}\" x2=\"{ox + pw}\" y2=\"{oy + plotH}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        // Draw depth guide lines on the floor (baseline) to show perspective
        for (int c = 0; c < catCount; c++)
        {
            var frontX = ox + (catCount > 1 ? (double)plotW * c / (catCount - 1) : plotW / 2.0);
            var backX = frontX + totalDepthX;
            var backY = oy + plotH + totalDepthY;
            sb.AppendLine($"        <line x1=\"{frontX:0.#}\" y1=\"{oy + plotH}\" x2=\"{backX:0.#}\" y2=\"{backY:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.3\"/>");
        }

        var stackBase = new double[catCount];

        // Draw back-to-front: back series first (farthest), front series last (nearest)
        for (int si = (stacked ? 0 : serCount - 1); stacked ? si < serCount : si >= 0; si += stacked ? 1 : -1)
        {
            var color = colors[si % colors.Count];
            var wallColor = AdjustColor(color, 0.6);
            var topColor = AdjustColor(color, 0.85);

            // Compute this series' lane position
            int lane = stacked ? 0 : si;
            // Front edge of this lane (start of wall)
            var laneDx = laneStep * lane * 0.7;
            var laneDy = -laneStep * lane * 0.5;
            // Back edge of this lane (end of wall = front + thickness)
            var nextDx = laneDx + laneThickness * 0.7;
            var nextDy = laneDy - laneThickness * 0.5;

            // Front edge points (data line at this lane's Z)
            var frontPts = new List<(double x, double y)>();
            // Back edge points (same data but shifted deeper)
            var backPts = new List<(double x, double y)>();

            for (int c = 0; c < catCount; c++)
            {
                var val = c < series[si].values.Length ? series[si].values[c] : 0;
                var baseVal = stacked ? stackBase[c] : 0;
                var topVal = baseVal + val;
                var dataH = (topVal / maxVal) * plotH;
                var baseH = (baseVal / maxVal) * plotH;

                var frontBaseX = ox + (catCount > 1 ? (double)plotW * c / (catCount - 1) : plotW / 2.0);

                var fx = frontBaseX + laneDx;
                var fy = oy + plotH - dataH + laneDy;
                frontPts.Add((fx, fy));

                var bx = frontBaseX + nextDx;
                var by = oy + plotH - dataH + nextDy;
                backPts.Add((bx, by));
            }

            if (frontPts.Count < 2) continue;

            // 1) Top ribbon: polygon connecting front data edge to back data edge (shows "roof" of the wall)
            var topPath = new StringBuilder("M ");
            foreach (var pt in frontPts) topPath.Append($"{pt.x:0.#},{pt.y:0.#} L ");
            for (int p = backPts.Count - 1; p >= 0; p--)
                topPath.Append($"{backPts[p].x:0.#},{backPts[p].y:0.#} L ");
            topPath.Length -= 2;
            topPath.Append(" Z");
            sb.AppendLine($"        <path d=\"{topPath}\" fill=\"{topColor}\" opacity=\"0.8\"/>");

            // 2) Front face: area from front baseline up to front data line
            var frontBaseY = oy + plotH + laneDy;
            var areaPath = new StringBuilder($"M {frontPts[0].x:0.#},{frontBaseY + (stacked ? -(stackBase[0] / maxVal) * plotH : 0):0.#} ");
            foreach (var pt in frontPts) areaPath.Append($"L {pt.x:0.#},{pt.y:0.#} ");
            areaPath.Append($"L {frontPts[^1].x:0.#},{frontBaseY + (stacked ? -(stackBase[catCount - 1] / maxVal) * plotH : 0):0.#} ");
            if (stacked)
            {
                for (int c = catCount - 1; c >= 0; c--)
                {
                    var baseX = ox + laneDx + (catCount > 1 ? (double)plotW * c / (catCount - 1) : plotW / 2.0);
                    var baseY2 = oy + plotH + laneDy - (stackBase[c] / maxVal) * plotH;
                    areaPath.Append($"L {baseX:0.#},{baseY2:0.#} ");
                }
            }
            areaPath.Append("Z");
            sb.AppendLine($"        <path d=\"{areaPath}\" fill=\"{color}\" opacity=\"0.9\"/>");

            // 3) Front edge line
            sb.AppendLine($"        <polyline points=\"{string.Join(" ", frontPts.Select(p => $"{p.x:0.#},{p.y:0.#}"))}\" fill=\"none\" stroke=\"{AdjustColor(color, 0.7)}\" stroke-width=\"1.5\"/>");

            // 4) Right-side wall (last category): connects front-right to back-right edge
            {
                var frX = frontPts[^1].x; var frY = frontPts[^1].y;
                var brX = backPts[^1].x; var brY = backPts[^1].y;
                var frBaseY2 = frontBaseY + (stacked ? -(stackBase[catCount - 1] / maxVal) * plotH : 0);
                var brBaseY = oy + plotH + nextDy + (stacked ? -(stackBase[catCount - 1] / maxVal) * plotH : 0);
                sb.AppendLine($"        <polygon points=\"{frX:0.#},{frY:0.#} {brX:0.#},{brY:0.#} {brX:0.#},{brBaseY:0.#} {frX:0.#},{frBaseY2:0.#}\" fill=\"{wallColor}\" opacity=\"0.8\"/>");
            }

            if (stacked)
            {
                for (int c = 0; c < catCount; c++)
                    stackBase[c] += c < series[si].values.Length ? series[si].values[c] : 0;
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)plotW * c / (catCount - 1) : plotW / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + plotH + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        // Value axis
        for (int t = 0; t <= 4; t++)
        {
            var val = maxVal * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + plotH - (double)plotH * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    // ==================== Trendline Regression Math ====================

    /// <summary>Least-squares linear regression: y = slope * x + intercept.</summary>
    private static (double slope, double intercept) FitLinear(double[] x, double[] y)
    {
        int n = x.Length;
        double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
        for (int i = 0; i < n; i++)
        {
            sumX += x[i]; sumY += y[i];
            sumXY += x[i] * y[i]; sumX2 += x[i] * x[i];
        }
        var denom = n * sumX2 - sumX * sumX;
        if (Math.Abs(denom) < 1e-15) return (0, sumY / n);
        var slope = (n * sumXY - sumX * sumY) / denom;
        var intercept = (sumY - slope * sumX) / n;
        return (slope, intercept);
    }

    /// <summary>Exponential fit: y = a * e^(b*x). Uses ln(y) linear regression.</summary>
    private static (double a, double b) FitExponential(double[] x, double[] y)
    {
        // Filter to positive y values only
        var validIdx = Enumerable.Range(0, y.Length).Where(i => y[i] > 0).ToArray();
        if (validIdx.Length < 2) return (double.NaN, double.NaN);
        var lnY = validIdx.Select(i => Math.Log(y[i])).ToArray();
        var xv = validIdx.Select(i => x[i]).ToArray();
        var (slope, intercept) = FitLinear(xv, lnY);
        return (Math.Exp(intercept), slope);
    }

    /// <summary>Logarithmic fit: y = a * ln(x) + b. Uses ln(x) linear regression.</summary>
    private static (double a, double b) FitLogarithmic(double[] x, double[] y)
    {
        var validIdx = Enumerable.Range(0, x.Length).Where(i => x[i] > 0).ToArray();
        if (validIdx.Length < 2) return (double.NaN, double.NaN);
        var lnX = validIdx.Select(i => Math.Log(x[i])).ToArray();
        var yv = validIdx.Select(i => y[i]).ToArray();
        var (slope, intercept) = FitLinear(lnX, yv);
        return (slope, intercept);
    }

    /// <summary>Power fit: y = a * x^b. Uses ln(x),ln(y) linear regression.</summary>
    private static (double a, double b) FitPower(double[] x, double[] y)
    {
        var validIdx = Enumerable.Range(0, x.Length).Where(i => x[i] > 0 && y[i] > 0).ToArray();
        if (validIdx.Length < 2) return (double.NaN, double.NaN);
        var lnX = validIdx.Select(i => Math.Log(x[i])).ToArray();
        var lnY = validIdx.Select(i => Math.Log(y[i])).ToArray();
        var (slope, intercept) = FitLinear(lnX, lnY);
        return (Math.Exp(intercept), slope);
    }

    /// <summary>Polynomial fit: y = c0 + c1*x + c2*x² + ... using normal equations.</summary>
    private static double[]? FitPolynomial(double[] x, double[] y, int order)
    {
        int n = x.Length;
        order = Math.Min(order, n - 1);
        if (order < 1) return null;
        int m = order + 1;

        // Build normal equations: (X^T X) c = X^T y
        var xtx = new double[m, m];
        var xty = new double[m];
        for (int i = 0; i < n; i++)
        {
            var xPow = new double[2 * order + 1];
            xPow[0] = 1;
            for (int p = 1; p <= 2 * order; p++) xPow[p] = xPow[p - 1] * x[i];
            for (int r = 0; r < m; r++)
            {
                for (int c = 0; c < m; c++) xtx[r, c] += xPow[r + c];
                xty[r] += xPow[r] * y[i];
            }
        }

        // Gaussian elimination with partial pivoting
        var aug = new double[m, m + 1];
        for (int r = 0; r < m; r++)
        {
            for (int c = 0; c < m; c++) aug[r, c] = xtx[r, c];
            aug[r, m] = xty[r];
        }
        for (int col = 0; col < m; col++)
        {
            int pivotRow = col;
            for (int r = col + 1; r < m; r++)
                if (Math.Abs(aug[r, col]) > Math.Abs(aug[pivotRow, col])) pivotRow = r;
            if (pivotRow != col)
                for (int c = 0; c <= m; c++) (aug[col, c], aug[pivotRow, c]) = (aug[pivotRow, c], aug[col, c]);
            if (Math.Abs(aug[col, col]) < 1e-15) return null;
            for (int r = col + 1; r < m; r++)
            {
                var factor = aug[r, col] / aug[col, col];
                for (int c = col; c <= m; c++) aug[r, c] -= factor * aug[col, c];
            }
        }
        // Back substitution
        var coeffs = new double[m];
        for (int r = m - 1; r >= 0; r--)
        {
            coeffs[r] = aug[r, m];
            for (int c = r + 1; c < m; c++) coeffs[r] -= aug[r, c] * coeffs[c];
            coeffs[r] /= aug[r, r];
        }
        return coeffs;
    }

    /// <summary>Compute R² (coefficient of determination).</summary>
    private static double ComputeRSquared(double[] x, double[] y, Func<double, double> fn)
    {
        var mean = y.Average();
        double ssTot = 0, ssRes = 0;
        for (int i = 0; i < y.Length; i++)
        {
            ssTot += (y[i] - mean) * (y[i] - mean);
            var predicted = fn(x[i]);
            ssRes += (y[i] - predicted) * (y[i] - predicted);
        }
        return ssTot > 0 ? 1 - ssRes / ssTot : 0;
    }
}
