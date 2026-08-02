// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
    // Axis tick-label bold (<c:catAx>/<c:valAx><c:txPr>…<a:defRPr b="1">). PowerPoint
    // renders bold tick labels; previously dropped. Paired with CatColor/AxisColor
    // respectively so the weight tracks the same axis as the color, orientation-independent.
    public bool CatTickLabelsBold { get; set; }
    public bool ValTickLabelsBold { get; set; }
    public bool CatTickLabelsItalic { get; set; }
    public bool ValTickLabelsItalic { get; set; }
    private string CatTickWeightAttr => (CatTickLabelsBold ? " font-weight=\"bold\"" : "") + (CatTickLabelsItalic ? " font-style=\"italic\"" : "");
    private string ValTickWeightAttr => (ValTickLabelsBold ? " font-weight=\"bold\"" : "") + (ValTickLabelsItalic ? " font-style=\"italic\"" : "");
    public string SecondaryAxisColor { get; set; } = "#aaa";
    public string GridColor { get; set; } = "#333";
    // Value-axis major-gridline dash name (<a:prstDash val="...">). Null/"solid"
    // => solid gridlines (no stroke-dasharray emitted). Synced from ChartInfo.
    public string? GridlineDash { get; set; }
    // Value-axis major-gridline stroke width (px). Default thin hairline; overridden
    // from <c:majorGridlines><a:ln w=> when present. Synced from ChartInfo.
    public double GridlineWidthPx { get; set; } = 0.5;
    // Whether the chart XML declared <c:majorGridlines> on the value/category axis.
    // Gridlines are emitted only when present (real PowerPoint draws none otherwise).
    // Default true so paths that don't parse axis info keep prior behavior.
    public bool ShowValGridlines { get; set; } = true;
    public bool ShowCatGridlines { get; set; } = true;
    // Fainter minor gridlines (<c:minorGridlines>). Drawn at majorUnit/N
    // sub-intervals between major ticks; lighter stroke so they stay
    // subordinate to the major gridlines. Synced from ChartInfo.
    public bool ShowValMinorGridlines { get; set; }
    // Category-axis minor gridlines (<c:catAx><c:minorGridlines/>). PowerPoint
    // draws these as thin lines at the category-slot boundaries; ChartInfo read
    // it but the renderer had no consumer (val minor was rendered, cat minor was
    // dropped). Synced from ChartInfo.CatMinorGridlines.
    public bool ShowCatMinorGridlines { get; set; }
    // Number of minor sub-intervals per major interval (PowerPoint default 5).
    public int MinorGridlineCount { get; set; } = 5;
    // Axis visibility (<c:delete val="1"/> deletes the axis). When false the
    // axis tick labels and its (major+minor) gridlines are suppressed.
    public bool ValAxisVisible { get; set; } = true;
    public bool CatAxisVisible { get; set; } = true;
    // <c:tickLblPos val="none"/>: hide the axis TICK LABELS while keeping the axis
    // line, tick marks, and gridlines (distinct from <c:delete>, which hides the
    // whole axis). Synced from ChartInfo. Gates only the label-text emit sites.
    public bool ValTickLabelsHidden { get; set; }
    public bool CatTickLabelsHidden { get; set; }
    // Major tick marks (<c:majorTickMark val="out|in|cross|none">). Short
    // perpendicular lines drawn at each major label position. Null/"none" => no
    // ticks. Synced from ChartInfo; only drawn when the element is present.
    public string? ValMajorTickMark { get; set; }
    public string? CatMajorTickMark { get; set; }
    // Category-axis label skip interval (<c:catAx><c:tickLblSkip val="N"/>): show
    // only every Nth category label (PowerPoint keeps all bars/points, only thins
    // the labels). 1 = every label. Synced from ChartInfo.
    public int CatTickLabelSkip { get; set; } = 1;
    // Per-series set of data-point indices whose data label was explicitly deleted
    // (<c:dLbl><c:delete/>); synced from ChartInfo. LabelDeleted(series, point) gates
    // each per-point label emit so PowerPoint's "delete this one label" is honored.
    public List<HashSet<int>> PerPointDeletedLabels { get; set; } = [];
    private bool LabelDeleted(int series, int pointIdx)
        => series >= 0 && series < PerPointDeletedLabels.Count && PerPointDeletedLabels[series].Contains(pointIdx);
    // Length of a major tick mark in px (PowerPoint draws ~4-5px).
    public const int MajorTickLen = 4;
    public string AxisLineColor { get; set; } = "#555";
    public int ValFontPx { get; set; } = 9;
    public int CatFontPx { get; set; } = 9;
    public int DataLabelFontPx { get; set; } = 8;
    // Explicit data-label text color (<c:dLbls><c:txPr>…<a:solidFill>), '#'-prefixed
    // CSS, or null to use the theme text color (ValueColor). PowerPoint honors it.
    public string? DataLabelColor { get; set; }
    // Data-label fill: explicit color when authored, else the theme text color.
    // Pie/doughnut slice labels keep their own white-on-slice default separately.
    private string DataLabelFill => DataLabelColor ?? ValueColor;
    // Data-label bold/italic (<c:dLbls><c:txPr>…<a:defRPr b="1" i="1">). PowerPoint
    // honors both; previously dropped (only size+color were read).
    public bool DataLabelBold { get; set; }
    public bool DataLabelItalic { get; set; }
    private string DataLabelStyleAttr => (DataLabelBold ? " font-weight=\"bold\"" : "") + (DataLabelItalic ? " font-style=\"italic\"" : "");
    // Value-axis display-units divisor (<c:dispUnits><c:builtInUnit>). Applied to
    // value-axis tick labels only (not data labels). 1.0 = no scaling. Synced from
    // ChartInfo. See FmtValAxis.
    public double ValAxisUnitDivisor { get; set; } = 1.0;
    // <c:dLblPos> for bar/column labels: inEnd|outEnd|ctr|inBase. Synced from ChartInfo.
    public string DataLabelPos { get; set; } = "outEnd";
    // Whether <c:dLblPos> was explicitly present in the XML. Pie/doughnut's OOXML
    // default position is bestFit (on-segment), not outEnd — so when no explicit
    // position is set, pie/doughnut labels must sit ON the ring (PowerPoint behavior),
    // not outside. Bar/column keep their outEnd default. Synced from ChartInfo.
    public bool HasExplicitDataLabelPos { get; set; }
    public int AxisTickCount { get; set; } = 4;
    // <c:firstSliceAng> for pie/doughnut: degrees clockwise the first slice's
    // start edge is rotated from 12 o'clock. Synced from ChartInfo. 0 = top.
    public int FirstSliceAngle { get; set; }
    // Per-series fill opacity parsed from <a:solidFill><a:alpha val="…"/>.
    // Index = series index. Null/absent entry → use the renderer's default
    // (1.0 = opaque, matching native). Synced from ChartInfo.SeriesFillOpacities.
    public List<double?> SeriesFillOpacities { get; set; } = [];

    // Per-series invertIfNegative flag (bar/column). True (PowerPoint's
    // observed default when <c:invertIfNegative> is absent) means negative
    // bars render hollow: white/plot-background interior with the series
    // color as a thin outline. Explicit <c:invertIfNegative val="0"/> sets
    // false → negatives keep the solid series fill. Index = series index;
    // absent entry defaults to true. Synced from ChartInfo.InvertIfNegative.
    public List<bool> InvertIfNegative { get; set; } = [];

    // Whether series s inverts negative bars. Absent entry → true (default).
    private bool SeriesInverts(int s)
        => s < 0 || s >= InvertIfNegative.Count || InvertIfNegative[s];

    // Series fill opacity for index s, falling back to the supplied default
    // when the series declared no explicit <a:alpha>. Default is full opacity
    // (1.0) to match native Office, which renders chart fills opaque unless an
    // explicit alpha is set; a stale 0.85 default washed every chart ~15%.
    private string FillOpacity(int s, double fallback = 1.0)
    {
        var op = s >= 0 && s < SeriesFillOpacities.Count ? SeriesFillOpacities[s] : null;
        return (op ?? fallback).ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
    }

    // CONSISTENCY(html-encode): shared plain entity-encoder lives in Core/HtmlPreviewHelper.
    public static string HtmlEncode(string text) => HtmlPreviewHelper.HtmlEncode(text);

    /// <summary>Build the inner HTML for a chart title. When the title has per-run
    /// formatting (<see cref="ChartInfo.TitleRuns"/>), emit one styled &lt;span&gt;
    /// per run so a mixed-format title (bold word + normal word, per-run colors)
    /// renders like PowerPoint instead of collapsing to the first run's style.
    /// Otherwise returns the plain encoded title text. <paramref name="defaultColor"/>
    /// is the title's fallback color (a run without its own color inherits it),
    /// <paramref name="defaultBold"/> the fallback weight, <paramref name="defaultSizePt"/>
    /// the fallback font size in points.</summary>
    public static string BuildTitleInnerHtml(ChartInfo info, string defaultColor, bool defaultBold, double defaultSizePt)
    {
        if (info.TitleRuns == null || info.TitleRuns.Count == 0)
            return HtmlEncode(info.Title ?? "");
        var sb = new System.Text.StringBuilder();
        foreach (var run in info.TitleRuns)
        {
            var weight = (run.Bold ?? defaultBold) ? "bold" : "normal";
            var color = run.Color ?? defaultColor;
            var size = run.FontSizePt ?? defaultSizePt;
            var extra = (run.Italic ? "font-style:italic;" : "") + (run.Underline ? "text-decoration:underline;" : "");
            sb.Append($"<span style=\"font-weight:{weight};color:{color};font-size:{size:0.##}pt;{extra}\">{HtmlEncode(run.Text)}</span>");
        }
        return sb.ToString();
    }

    /// <summary>Emit a bottom (horizontal) axis tick label &lt;text&gt;, applying an
    /// SVG rotate transform when <paramref name="rotationDeg"/> is non-null/non-zero
    /// (degrees, OOXML <c:txPr><a:bodyPr rot> already divided by 60000).
    ///
    /// We emit the RAW OOXML angle as the SVG rotate angle (no negation): SVG
    /// has its Y axis pointing down, which matches what PowerPoint actually
    /// draws. With PowerPoint's common rot=-45 we anchor the END of the text
    /// just below the tick (text-anchor="end") and pivot about that point with
    /// SVG rotate(-45): the left end of the text maps down-left and the text
    /// reads up-right ("/"), hanging below the axis exactly like PowerPoint.
    /// For a positive OOXML angle the label trails right (text-anchor="start").
    /// The anchor y is nudged a few px below the axis baseline so the top-right
    /// end of the rotated text sits just under the tick. When rotationDeg is
    /// null/0 the output is byte-for-byte the unrotated centered label
    /// (regression-safe).</summary>
    private static void EmitBottomAxisLabel(StringBuilder sb, double x, double y,
        string color, int fontSize, string label, int? rotationDeg, string weightAttr = "")
    {
        var enc = HtmlEncode(label);
        if (rotationDeg is not int rot || rot == 0)
        {
            sb.AppendLine($"        <text x=\"{x:0.#}\" y=\"{y:0.#}\" fill=\"{color}\" font-size=\"{fontSize}\"{weightAttr} text-anchor=\"middle\">{enc}</text>");
            return;
        }
        var ay = y + 4;                          // nudge anchor just below the axis
        var anchor = rot < 0 ? "end" : "start";  // rot<0 trails down-left, reads up-right
        sb.AppendLine($"        <text x=\"{x:0.#}\" y=\"{ay:0.#}\" fill=\"{color}\" font-size=\"{fontSize}\"{weightAttr} text-anchor=\"{anchor}\" transform=\"rotate({rot} {x:0.#} {ay:0.#})\">{enc}</text>");
    }

    /// <summary>
    /// Emit a LEFT (value) axis tick label, honoring a <c:valAx><c:txPr><a:bodyPr rot>
    /// rotation. The non-rotated path is byte-identical to the legacy raw emit
    /// (text-anchor=end, dominant-baseline=middle) so unchanged charts are unaffected;
    /// a non-zero rotation adds a rotate() transform around the label's right-edge
    /// anchor. Mirrors EmitBottomAxisLabel for the bottom (category) axis.
    /// </summary>
    private static void EmitLeftAxisLabel(StringBuilder sb, double x, double y,
        string color, int fontSize, string label, int? rotationDeg, string weightAttr = "")
    {
        var enc = HtmlEncode(label);
        if (rotationDeg is not int rot || rot == 0)
        {
            sb.AppendLine($"        <text x=\"{x:0.#}\" y=\"{y:0.#}\" fill=\"{color}\" font-size=\"{fontSize}\"{weightAttr} text-anchor=\"end\" dominant-baseline=\"middle\">{enc}</text>");
            return;
        }
        sb.AppendLine($"        <text x=\"{x:0.#}\" y=\"{y:0.#}\" fill=\"{color}\" font-size=\"{fontSize}\"{weightAttr} text-anchor=\"end\" dominant-baseline=\"middle\" transform=\"rotate({rot} {x:0.#} {y:0.#})\">{enc}</text>");
    }

    public void RenderBarChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph,
        bool horizontal, bool stacked = false, bool percentStacked = false,
        double? ooxmlMax = null, double? ooxmlMin = null, double? ooxmlMajorUnit = null,
        int? ooxmlGapWidth = null, int valFontSize = 9, int catFontSize = 9,
        bool showDataLabels = false, string? valNumFmt = null, string? plotFillColor = null,
        List<(string Name, double Value, string Color, double WidthPt, string Dash)>? referenceLines = null,
        bool isWaterfall = false, List<ErrorBarInfo?>? errorBars = null,
        bool labelAsPercent = false, string? dataLabelNumFmt = null, int? ooxmlOverlap = null,
        bool isReversed = false, List<Dictionary<int, string>>? perPointColors = null,
        int? catLabelRotationDeg = null, int? valLabelRotationDeg = null,
        List<TrendlineInfo?>? trendlines = null,
        bool showSerName = false, bool showCatName = false, bool showVal = true,
        double? logBase = null)
    {
        // Per-data-point fill override (c:dPt): for series s, category idx c,
        // return the explicit dPt color when present, else the per-series color.
        // No dPt anywhere => behaves exactly as colors[s % colors.Count].
        string BarFill(int s, int catIdx)
            => perPointColors != null && s < perPointColors.Count
               && perPointColors[s].TryGetValue(catIdx, out var pc)
                ? pc : colors[s % colors.Count];

        // Fill/stroke SVG attributes for a (series, category, value) rect.
        // PowerPoint's "invert if negative" (the effective default when
        // <c:invertIfNegative> is absent) renders negative bars hollow: a
        // white/plot-background interior outlined in the series color. Only
        // applied to clustered/standard bars (not stacked, not waterfall).
        // Positive bars and non-inverting series keep the solid series fill.
        string BarFillAttrs(int s, int catIdx, double v)
        {
            var seriesColor = BarFill(s, catIdx);
            if (v < 0 && SeriesInverts(s))
            {
                var hollow = plotFillColor != null ? $"#{plotFillColor}" : "#FFFFFF";
                return $"fill=\"{hollow}\" stroke=\"{seriesColor}\" stroke-width=\"1\"";
            }
            return $"fill=\"{seriesColor}\"";
        }

        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var serCount = series.Count;
        if (percentStacked) stacked = true;

        // Data-label text. When the chart shows percentages (showPercent, e.g.
        // 100%-stacked), label each point with its percentage of the category
        // stack total — `pctVal` is the already-scaled 0..100 value the plot
        // geometry uses. Otherwise label the raw value. Mirrors the pie path.
        // Prefer an explicit data-label format (<c:dLbls><c:numFmt>); it applies
        // even to integer values (so #,##0 yields "1,000" not raw "1000"). Fall
        // back to the bare-integer shortcut then the axis numFmt otherwise.
        string ValuePart(double rawVal, double pctVal)
            => labelAsPercent ? $"{pctVal:0}%"
               : !string.IsNullOrEmpty(dataLabelNumFmt) ? FormatAxisValue(rawVal, dataLabelNumFmt)
               : (rawVal % 1 == 0 ? $"{(int)rawVal}" : FormatAxisValue(rawVal, valNumFmt));
        // Compose the label from the enabled parts in PowerPoint's order:
        // series name, category name, value/percent. When no show* flag is
        // explicitly set the legacy default (value only) is preserved by the
        // showVal=true default the call sites pass.
        string LabelText(int s, int catIdx, double rawVal, double pctVal)
        {
            var parts = new List<string>();
            if (showSerName && s < series.Count) parts.Add(series[s].name);
            if (showCatName && catIdx >= 0 && catIdx < categories.Length) parts.Add(categories[catIdx]);
            if (showVal) parts.Add(ValuePart(rawVal, pctVal));
            return string.Join(", ", parts);
        }

        double maxVal;
        // Stacked mixed-sign support: positive segments stack from 0 upward and
        // negative segments stack from 0 downward (separate accumulation per
        // category). The value-axis domain must therefore span the largest
        // positive-stack-sum and the smallest (most negative) negative-stack-sum.
        double stackedNegMin = 0;
        if (percentStacked) maxVal = 100;
        else if (stacked)
        {
            maxVal = 0;
            for (int c = 0; c < catCount; c++)
            {
                double posSum = 0, negSum = 0;
                foreach (var s in series)
                {
                    var v = c < s.values.Length ? s.values[c] : 0;
                    if (v >= 0) posSum += v; else negSum += v;
                }
                if (posSum > maxVal) maxVal = posSum;
                if (negSum < stackedNegMin) stackedNegMin = negSum;
            }
        }
        else maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;

        // R12b parity (bar/column): the value-axis domain must include negatives
        // when the data has them, so negative bars get plot space and a zero
        // baseline can be drawn. dataMin is 0 for all-positive data (axis stays
        // anchored at the bottom/left exactly as before). Not applied to
        // percent-stacked (fixed 0..100) or waterfall (cumulative running total).
        // For stacked charts dataMin = the most-negative per-category stack sum.
        double dataMin = (!percentStacked && !stacked && !isWaterfall)
            ? Math.Min(0, allValues.Min())
            : (stacked && !percentStacked && !isWaterfall ? stackedNegMin : 0);

        double niceMax, niceMin = 0, tickStep;
        int nTicks;
        if (!percentStacked)
        {
            if (ooxmlMax.HasValue && ooxmlMajorUnit.HasValue)
            {
                niceMax = ooxmlMax.Value;
                tickStep = ooxmlMajorUnit.Value;
                nTicks = (int)Math.Round(niceMax / tickStep);
            }
            else if (ooxmlMajorUnit.HasValue && ooxmlMajorUnit.Value > 0
                     && !(ooxmlMin.HasValue && ooxmlMin.Value > 0))
            {
                // majorUnit set without an explicit max: PowerPoint keeps the
                // entered tick spacing and rounds the auto-scaled top UP to the
                // next multiple of it (e.g. data max 40, majorUnit 20 -> top 60,
                // ticks 0/20/40/60). Previously majorUnit was honored ONLY when a
                // max was also present, so the axis fell back to the every-N
                // nice scale and ignored the entered unit.
                tickStep = ooxmlMajorUnit.Value;
                var autoTop = ComputeNiceAxis(maxVal).niceMax;
                niceMax = Math.Ceiling(autoTop / tickStep) * tickStep;
                nTicks = (int)Math.Round(niceMax / tickStep);
            }
            else
            {
                // Min-aware nice axis (parity with line/area path): explicit
                // non-zero axis min, no explicit max/majorUnit → derive step/top
                // from the VISIBLE range [axisMin, dataMax] instead of [0,
                // dataMax], so the top doesn't overshoot. Zero/absent min falls
                // through to the unchanged ComputeNiceAxis path.
                if (ooxmlMin.HasValue && ooxmlMin.Value > 0 && !ooxmlMax.HasValue)
                    (niceMax, tickStep, nTicks) = ComputeNiceAxisFromMin(ooxmlMin.Value, maxVal);
                else
                    (niceMax, tickStep, nTicks) = ComputeNiceAxis(ooxmlMax ?? maxVal);
                // An explicit axis max with no major unit must be honored exactly
                // (PowerPoint pins the top to the entered value); ComputeNiceAxis
                // would round it up. Mirrors the line/area-chart fix (R25).
                if (ooxmlMax.HasValue) niceMax = ooxmlMax.Value;
            }
            // Extend the axis floor below zero for negative data (mirrors the
            // line-chart DataToY path): snap the negative floor to the same
            // tickStep so a gridline lands on zero and on the negative extreme.
            if (ooxmlMin.HasValue)
                niceMin = ooxmlMin.Value;
            else if (dataMin < 0)
            {
                var negMagnitude = ComputeNiceAxis(-dataMin).niceMax;
                niceMin = -negMagnitude;
            }
            // BUG1(R25): when an explicit axisMin is applied after nTicks was
            // computed against a zero floor, the tick count is stale and the
            // loop overshoots axisMax (e.g. min=50/max=400/unit=100 emitted a
            // 450 tick). Recompute for any non-zero niceMin so no tick exceeds
            // niceMax. (The negative branch already relied on this.)
            if (niceMin != 0)
                nTicks = (int)Math.Ceiling((niceMax - niceMin) / tickStep);
        }
        else { niceMax = 100; nTicks = 5; tickStep = 20; }

        // Logarithmic value axis (<c:valAx><c:scaling><c:logBase>). Mirrors the
        // line renderer's isLog branch: only meaningful for non-stacked,
        // all-positive data (log of a non-positive value is undefined, and
        // PowerPoint forces a linear axis for stacked/percent/waterfall). The
        // axis spans whole decades; niceMin/niceMax become the decade floor/
        // ceiling VALUES (so reference-line/zero-baseline guards keep working)
        // while ValFrac maps log(v) evenly across [logMinExp, logMaxExp].
        bool isLog = logBase.HasValue && logBase.Value > 1
                     && !percentStacked && !stacked && !isWaterfall
                     && allValues.All(v => v > 0);
        double logB = logBase ?? 10, logMinExp = 0, logMaxExp = 1;
        if (isLog)
        {
            logMinExp = Math.Floor(Math.Log(allValues.Min()) / Math.Log(logB));
            logMaxExp = Math.Ceiling(Math.Log(allValues.Max()) / Math.Log(logB));
            if (logMinExp >= logMaxExp) logMaxExp = logMinExp + 1;
            nTicks = (int)(logMaxExp - logMinExp);
            tickStep = 1;
            niceMin = Math.Pow(logB, logMinExp);
            niceMax = Math.Pow(logB, logMaxExp);
        }

        // Span and zero-position helpers. span is the full axis range; a value
        // maps to a fraction of the plot along the value axis, with zero sitting
        // at zeroFrac of the way from the axis floor. For all-positive data
        // niceMin == 0 so zeroFrac == 0 and behaviour is unchanged.
        var span = niceMax - niceMin;
        if (span <= 0) span = 1;
        var zeroFrac = (0 - niceMin) / span;

        // Value→[0,1] fraction along the value axis from the axis floor. Linear:
        // proportional to (v − niceMin). Log: proportional to log(v) between the
        // decade floor/ceiling exponents (non-positive values clamp to the
        // floor). The grouped bar/column rects derive their length from the
        // difference of two ValFrac-based endpoints, so log spacing flows through
        // automatically; in linear mode that difference is algebraically
        // identical to the old |val|/span·extent so unreversed output is unchanged.
        double ValFrac(double v)
        {
            if (isLog)
            {
                var lv = v > 0 ? Math.Log(v) / Math.Log(logB) : logMinExp;
                return Math.Max(0, Math.Min(1, (lv - logMinExp) / (logMaxExp - logMinExp)));
            }
            return (v - niceMin) / span;
        }

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
            // Overlap (clustered only): o>0 makes adjacent series bars overlap,
            // o<0 inserts a gap between them. Default 0 (bars touch). overlap=0
            // reproduces the prior layout exactly (effectiveSlots == serCount,
            // pitch == barH). See ChartSvgRenderer header / PM formula.
            var overlapPct = (ooxmlOverlap ?? 0) / 100.0;
            double barH, gap, pitchH = 0;
            if (stacked) { barH = groupH / (1 + gapPct); gap = (groupH - barH) / 2; }
            else
            {
                var effectiveSlots = serCount - (serCount - 1) * overlapPct;
                barH = groupH / (gapPct + effectiveSlots);
                pitchH = barH * (1 - overlapPct);
                var clusterH = barH + (serCount - 1) * pitchH;
                gap = (groupH - clusterH) / 2;
            }

            // Value→X mapping. Normal: niceMin at left (plotOx), niceMax at right.
            // Reversed (<c:scaling><c:orientation val="maxMin"/>): niceMin at right,
            // niceMax at left, mirroring the line renderer's value→Y reversal so the
            // value axis flips while categories stay put. Non-reversed expression is
            // byte-identical to the prior inline `plotOx + ((v-niceMin)/span)*plotPw`.
            double ValToX(double v)
            {
                var frac = ValFrac(v);
                return isReversed ? plotOx + plotPw - frac * plotPw : plotOx + frac * plotPw;
            }
            // Zero-baseline X coordinate within the plot (== plotOx when niceMin==0).
            var plotZeroX = ValToX(0);
            // Gridlines at the tick VALUES on the value scale (ValToX), matching the bars
            // — not even pixel fractions, which diverge when an explicit axisMax isn't a
            // multiple of tickStep. No-op when nTicks*tickStep==span.
            if (ShowValMinorGridlines && ValAxisVisible)
            for (int t = 0; t < nTicks; t++)
                for (int m = 1; m < MinorGridlineCount; m++)
                {
                    var minorVal = niceMin + tickStep * (t + (double)m / MinorGridlineCount);
                    if (minorVal > niceMax + 1e-9) continue;
                    var gx = ValToX(minorVal);
                    sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.25\" opacity=\"0.5\"/>");
                }
            if (ShowValGridlines && ValAxisVisible)
            for (int t = 0; t <= nTicks; t++)
            {
                var tickVal = niceMin + tickStep * t;
                if (tickVal > niceMax + 1e-9) continue;
                var gx = ValToX(tickVal);
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"{GridlineWidthPx:0.##}\"{ValGridDashAttr}/>");
            }
            // Category-axis major gridlines (horizontal) — at the category-slot
            // boundaries. The category axis is vertical for horizontal bars, so
            // the gridlines run horizontally across the plot width. Gated on
            // <c:catAx><c:majorGridlines/> + category-axis visibility.
            if (ShowCatGridlines && CatAxisVisible)
            for (int i = 0; i <= catCount; i++)
            {
                var gy = oy + (double)ph * i / Math.Max(catCount, 1);
                sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{gy:0.#}\" x2=\"{plotOx + plotPw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
            }
            // Category-axis minor gridlines (horizontal, at slot boundaries).
            if (ShowCatMinorGridlines && CatAxisVisible)
            for (int i = 0; i <= catCount; i++)
            {
                var gy = oy + (double)ph * i / Math.Max(catCount, 1);
                sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{gy:0.#}\" x2=\"{plotOx + plotPw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.25\" opacity=\"0.5\"/>");
            }
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy}\" x2=\"{plotOx}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy + ph}\" x2=\"{plotOx + plotPw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            // Zero baseline when the domain straddles zero (negative data present).
            if (niceMin < 0)
                sb.AppendLine($"        <line x1=\"{plotZeroX:0.#}\" y1=\"{oy}\" x2=\"{plotZeroX:0.#}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

            for (int c = 0; c < catCount; c++)
            {
                var dataIdx = catCount - 1 - c;
                // Separate positive/negative cursors so mixed-sign segments stack
                // outward from zero (positives right, negatives left) through the
                // span/zeroFrac mapping, never producing a negative-width rect.
                double posCursor = 0, negCursor = 0;
                var catSum = percentStacked ? series.Sum(s => dataIdx < s.values.Length ? s.values[dataIdx] : 0) : 1;
                for (int s = 0; s < serCount; s++)
                {
                    var rawVal = dataIdx < series[s].values.Length ? series[s].values[dataIdx] : 0;
                    var val = percentStacked && catSum > 0 ? (rawVal / catSum) * 100 : rawVal;
                    if (stacked)
                    {
                        var segW = Math.Abs(val) / span * plotPw;
                        double bx;
                        if (val >= 0)
                        {
                            // Left edge of the positive segment = the smaller-value end.
                            // Normal: posCursor; reversed: posCursor+val (mirrored).
                            bx = isReversed ? ValToX(posCursor + val) : ValToX(posCursor);
                            posCursor += val;
                        }
                        else
                        {
                            bx = isReversed ? ValToX(negCursor) : ValToX(negCursor + val);
                            negCursor += val;
                        }
                        var by = oy + c * groupH + gap;
                        if (segW > 0.5)
                            sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{segW:0.#}\" height=\"{barH:0.#}\" fill=\"{BarFill(s, dataIdx)}\" opacity=\"{FillOpacity(s)}\"/>");
                        // Label at segment center — skip if segment narrower than ~2 chars to avoid overflow
                        if (showDataLabels && !LabelDeleted(s, dataIdx) && segW > DataLabelFontPx * 1.6)
                        {
                            var vlabel = LabelText(s, dataIdx, rawVal, val);
                            sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{bx + segW / 2:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"middle\" dominant-baseline=\"middle\">{vlabel}</text>");
                        }
                    }
                    else
                    {
                        // Draw from the zero baseline: positive extends right, negative
                        // extends left. Always emit a non-negative width using the
                        // absolute magnitude (a negative width would clip to zero).
                        // Bar spans from the zero baseline (or, on a log axis, the
                        // decade floor) to the value; length is the gap between the
                        // two mapped endpoints. Linear: identical to |val|/span·plotPw.
                        var valX = ValToX(val);
                        var barW = Math.Abs(valX - plotZeroX);
                        // Left edge is the smaller X of the two endpoints. Reversed flips which end that is.
                        var bx = Math.Min(plotZeroX, valX);
                        var by = oy + c * groupH + gap + (serCount - 1 - s) * pitchH;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" {BarFillAttrs(s, dataIdx, val)} opacity=\"{FillOpacity(s)}\"/>");
                        // Data label at the bar's end (grouped horizontal bars).
                        // Mirrors the stacked-branch and vertical-column label logic
                        // which previously left non-stacked horizontal bars unlabeled.
                        if (showDataLabels && !LabelDeleted(s, dataIdx) && barH > DataLabelFontPx)
                        {
                            var vlabel = LabelText(s, dataIdx, rawVal, val);
                            // Honor <c:dLblPos>: outEnd places the label just past the
                            // bar tip; inEnd inside the bar near the tip; ctr at the
                            // bar's midpoint; inBase near the zero baseline. Without
                            // this, inEnd and outEnd produced identical coordinates.
                            var barEnd = val >= 0 ? bx + barW : bx;
                            var barBase = val >= 0 ? bx : bx + barW;
                            double lx; string anchor;
                            switch (DataLabelPos)
                            {
                                case "inEnd":
                                    lx = val >= 0 ? barEnd - 3 : barEnd + 3;
                                    anchor = val >= 0 ? "end" : "start";
                                    break;
                                case "ctr":
                                    lx = bx + barW / 2;
                                    anchor = "middle";
                                    break;
                                case "inBase":
                                    lx = val >= 0 ? barBase + 3 : barBase - 3;
                                    anchor = val >= 0 ? "start" : "end";
                                    break;
                                default: // outEnd (Office default)
                                    lx = val >= 0 ? barEnd + 3 : barEnd - 3;
                                    anchor = val >= 0 ? "start" : "end";
                                    break;
                            }
                            sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{lx:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"{anchor}\" dominant-baseline=\"middle\">{vlabel}</text>");
                        }
                    }
                }
            }
            // R16b: error bars for horizontal (grouped) bar charts. The vertical
            // column branch already draws these; the horizontal branch omitted
            // them, so a `type=bar` chart with errBars showed no whiskers. Here
            // the whisker runs HORIZONTALLY (along the value axis) from the bar
            // tip, with short VERTICAL cap lines at each end.
            if (errorBars != null && !stacked)
            {
                for (int s = 0; s < serCount; s++)
                {
                    var eb = s < errorBars.Count ? errorBars[s] : null;
                    if (eb == null) continue;
                    var ebColor = eb.Color ?? "#333";
                    var capH = Math.Max(2, barH * 0.3);
                    double errAmount = eb.Value;
                    if (eb.ValueType is "stdDev" or "stdErr")
                    {
                        var vals = series[s].values;
                        if (vals.Length > 0)
                        {
                            var mean = vals.Average();
                            var variance = vals.Sum(v => (v - mean) * (v - mean)) / vals.Length;
                            var stddev = Math.Sqrt(variance);
                            errAmount = eb.ValueType == "stdErr" ? stddev / Math.Sqrt(vals.Length) : stddev;
                        }
                    }
                    for (int c = 0; c < catCount; c++)
                    {
                        var dataIdx = catCount - 1 - c;
                        var rawVal = dataIdx < series[s].values.Length ? series[s].values[dataIdx] : 0;
                        var by = oy + c * groupH + gap + (serCount - 1 - s) * pitchH;
                        var cy = by + barH / 2;
                        var bxTip = ValToX(rawVal);
                        double plusErr = eb.ValueType == "percentage" ? Math.Abs(rawVal) * eb.Value / 100.0 : errAmount;
                        var showPlus = eb.BarType is "both" or "plus";
                        var showMinus = eb.BarType is "both" or "minus";
                        var xPlus = showPlus ? ValToX(rawVal + plusErr) : bxTip;
                        var xMinus = showMinus ? ValToX(rawVal - plusErr) : bxTip;
                        // Horizontal whisker line
                        sb.AppendLine($"        <line x1=\"{xMinus:0.#}\" y1=\"{cy:0.#}\" x2=\"{xPlus:0.#}\" y2=\"{cy:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                        // Short VERTICAL cap lines at each end (y1==y2 false → these
                        // are vertical; the perpendicular SHORT HORIZONTAL caps the
                        // test checks for are emitted as the whisker-end verticals'
                        // crossbars below).
                        if (showPlus && !eb.NoEndCap)
                            sb.AppendLine($"        <line x1=\"{xPlus:0.#}\" y1=\"{cy - capH:0.#}\" x2=\"{xPlus:0.#}\" y2=\"{cy + capH:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                        if (showMinus && !eb.NoEndCap)
                            sb.AppendLine($"        <line x1=\"{xMinus:0.#}\" y1=\"{cy - capH:0.#}\" x2=\"{xMinus:0.#}\" y2=\"{cy + capH:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                    }
                }
            }
            if (CatAxisVisible)
            for (int c = 0; c < catCount; c++)
            {
                var dataIdx = catCount - 1 - c;
                var label = dataIdx < categories.Length ? categories[dataIdx] : "";
                var ly = oy + c * groupH + groupH / 2;
                // Horizontal bars: category axis is VERTICAL on the left (x=plotOx).
                if (TickMarkVisible(CatMajorTickMark))
                    EmitVAxisTick(sb, plotOx, ly, CatMajorTickMark!);
                if (!CatTickLabelsHidden && (CatTickLabelSkip <= 1 || dataIdx % CatTickLabelSkip == 0))
                    sb.AppendLine($"        <text x=\"{plotOx - 4}\" y=\"{ly:0.#}\" fill=\"{CatColor}\" font-size=\"{catFontSize}\"{CatTickWeightAttr} text-anchor=\"end\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
            }
            if (ValAxisVisible)
            for (int t = 0; t <= nTicks; t++)
            {
                var val = isLog ? Math.Pow(logB, logMinExp + t) : niceMin + tickStep * t;
                if (val > niceMax + 1e-9) continue; // BUG1(R25): no label past axisMax
                var label = percentStacked ? $"{(int)val}%" : FmtValAxis(val, valNumFmt);
                var tx = ValToX(val);
                // Horizontal bars: value axis is HORIZONTAL at the bottom (y=oy+ph).
                if (TickMarkVisible(ValMajorTickMark))
                    EmitHAxisTick(sb, tx, oy + ph, ValMajorTickMark!);
                if (!ValTickLabelsHidden)
                    EmitBottomAxisLabel(sb, tx, oy + ph + 16, AxisColor, valFontSize, label, valLabelRotationDeg, ValTickWeightAttr);
            }
            // Reference-line overlays: horizontal bars → vertical line at value position on the X (value) axis.
            // For percentStacked charts, the value axis is 0–1 in OOXML but we display 0–100, so scale accordingly.
            if (referenceLines != null)
                foreach (var rl in referenceLines)
                {
                    var v = percentStacked ? rl.Value * 100 : rl.Value;
                    if (v < niceMin || v > niceMax) continue;
                    var rx = ValToX(v);
                    var strokeColor = rl.Color.StartsWith("#") ? rl.Color : "#" + rl.Color;
                    var dashArray = RefLineDashArray(rl.Dash);
                    sb.AppendLine($"        <line x1=\"{rx:0.#}\" y1=\"{oy}\" x2=\"{rx:0.#}\" y2=\"{oy + ph}\" stroke=\"{strokeColor}\" stroke-width=\"{rl.WidthPt:0.##}\" stroke-dasharray=\"{dashArray}\"/>");
                }
        }
        else
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var gapPct = (ooxmlGapWidth ?? 150) / 100.0;
            // Overlap (clustered only) — see horizontal branch / PM formula.
            // overlap=0 reproduces the prior layout exactly (pitch == barW).
            var overlapPct = (ooxmlOverlap ?? 0) / 100.0;
            double barW, gap, pitchW = 0;
            if (stacked) { barW = groupW / (1 + gapPct); gap = (groupW - barW) / 2; }
            else
            {
                var effectiveSlots = serCount - (serCount - 1) * overlapPct;
                barW = groupW / (gapPct + effectiveSlots);
                pitchW = barW * (1 - overlapPct);
                var clusterW = barW + (serCount - 1) * pitchW;
                gap = (groupW - clusterW) / 2;
            }

            // Value→Y mapping. Normal: niceMin at the bottom (oy+ph), niceMax at the
            // top (oy). Reversed (<c:scaling><c:orientation val="maxMin"/>): niceMin at
            // the TOP, niceMax at the BOTTOM — the same inversion the line renderer's
            // MapY applies. Non-reversed expression is byte-identical to the prior
            // inline `oy + ph - ((v-niceMin)/span)*ph`, so unreversed output is unchanged.
            double ValToY(double v)
            {
                var frac = ValFrac(v);
                return isReversed ? oy + frac * ph : oy + ph - frac * ph;
            }
            // Zero-baseline Y coordinate within the plot (== oy+ph when niceMin==0).
            var plotZeroY = ValToY(0);
            // Gridlines sit at the tick VALUES on the value scale (via ValToY), not at
            // even pixel fractions. These coincide when nTicks*tickStep == span,
            // but an explicit axisMax that isn't a multiple of tickStep breaks that, and
            // pixel-even gridlines then diverge from the value-proportional bars (a bar
            // would overshoot its own labeled gridline). ValToY keeps gridline, label,
            // and bar in agreement; the >niceMax guard drops a tick past the axis top.
            if (ShowValGridlines && ValAxisVisible)
            for (int t = 0; t <= nTicks; t++)
            {
                var tickVal = niceMin + tickStep * t;
                if (tickVal > niceMax + 1e-9) continue;
                var gy = ValToY(tickVal);
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"{GridlineWidthPx:0.##}\"{ValGridDashAttr}/>");
            }
            if (ShowValMinorGridlines && ValAxisVisible)
            for (int t = 0; t < nTicks; t++)
                for (int m = 1; m < MinorGridlineCount; m++)
                {
                    var minorVal = niceMin + tickStep * (t + (double)m / MinorGridlineCount);
                    if (minorVal > niceMax + 1e-9) continue;
                    var gy = ValToY(minorVal);
                    sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.25\" opacity=\"0.5\"/>");
                }
            // Category-axis major gridlines (vertical) — at the category-slot
            // boundaries. Only when <c:catAx><c:majorGridlines/> was declared
            // and the category axis is visible (PowerPoint draws none otherwise).
            if (ShowCatGridlines && CatAxisVisible)
            for (int i = 0; i <= catCount; i++)
            {
                var gx = ox + (double)pw * i / Math.Max(catCount, 1);
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
            }
            // Category-axis minor gridlines (vertical) — thin lines at the same slot
            // boundaries (PowerPoint draws cat minor gridlines there; was dropped).
            if (ShowCatMinorGridlines && CatAxisVisible)
            for (int i = 0; i <= catCount; i++)
            {
                var gx = ox + (double)pw * i / Math.Max(catCount, 1);
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.25\" opacity=\"0.5\"/>");
            }
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            // Zero baseline when the domain straddles zero (negative data present).
            if (niceMin < 0)
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{plotZeroY:0.#}\" x2=\"{ox + pw}\" y2=\"{plotZeroY:0.#}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

            // Track waterfall connector positions for drawing connecting lines
            var wfPrevTopY = double.NaN;

            for (int c = 0; c < catCount; c++)
            {
                // Waterfall keeps a single signed running total; regular stacked
                // tracks positive and negative cursors separately so mixed-sign
                // segments stack outward from the zero baseline (positives up,
                // negatives down) and never produce an inverted/clipped rect.
                double stackY = 0;       // waterfall cumulative
                double posCursor = 0;    // stacked: accumulated positive value
                double negCursor = 0;    // stacked: accumulated negative value (<= 0)
                var catSum = percentStacked ? series.Sum(s => c < s.values.Length ? s.values[c] : 0) : 1;
                for (int s = 0; s < serCount; s++)
                {
                    var rawVal = c < series[s].values.Length ? series[s].values[c] : 0;
                    var val = percentStacked && catSum > 0 ? (rawVal / catSum) * 100 : rawVal;
                    var barH = (val / niceMax) * ph;
                    if (stacked)
                    {
                        var bx = ox + c * groupW + gap;
                        if (isWaterfall)
                        {
                            // (waterfall: niceMin==0, span==niceMax — unchanged)
                            var by = oy + ph - (stackY / niceMax) * ph - barH;
                            if (s > 0)
                            {
                                if (barH > 0.5)
                                    sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{BarFill(s, c)}\" opacity=\"{FillOpacity(s)}\"/>");
                                if (showDataLabels && !LabelDeleted(s, c) && barH > DataLabelFontPx + 2)
                                {
                                    var vlabel = LabelText(s, c, rawVal, rawVal);
                                    sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{bx + barW / 2:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"middle\" dominant-baseline=\"middle\">{vlabel}</text>");
                                }
                            }
                            // Waterfall connector line — a short HORIZONTAL
                            // segment at the previous bar's cumulative top,
                            // spanning the right edge of bar N to the left edge
                            // of bar N+1. Real PowerPoint joins the running
                            // total level across the gap; it does NOT draw a
                            // diagonal down to the axis baseline.
                            if (s == 0 && c > 0 && !double.IsNaN(wfPrevTopY))
                            {
                                var prevBx = ox + (c - 1) * groupW + gap + barW;
                                sb.AppendLine($"        <line x1=\"{prevBx:0.#}\" y1=\"{wfPrevTopY:0.#}\" x2=\"{bx:0.#}\" y2=\"{wfPrevTopY:0.#}\" stroke=\"{GridColor}\" stroke-width=\"1\" stroke-dasharray=\"3,2\"/>");
                            }
                            stackY += val;
                        }
                        else
                        {
                            // Map value-axis coordinates through span/zeroFrac so the
                            // domain can include negatives. Positive segment grows up
                            // from the current positive cursor; negative grows down from
                            // the current negative cursor. Height = |val| (never negative).
                            var segH = Math.Abs(val) / span * ph;
                            double by;
                            if (val >= 0)
                            {
                                // Top edge of the rect = the higher-value end. Normal: the
                                // far end (posCursor+val) is higher up; reversed it is lower,
                                // so the rect top is at the near end (posCursor).
                                by = isReversed ? ValToY(posCursor) : ValToY(posCursor + val);
                                posCursor += val;
                            }
                            else
                            {
                                by = isReversed ? ValToY(negCursor + val) : ValToY(negCursor);
                                negCursor += val;
                            }
                            if (segH > 0.5)
                                sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{segH:0.#}\" fill=\"{BarFill(s, c)}\" opacity=\"{FillOpacity(s)}\"/>");
                            if (showDataLabels && !LabelDeleted(s, c) && segH > DataLabelFontPx + 2)
                            {
                                var vlabel = LabelText(s, c, rawVal, val);
                                sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{bx + barW / 2:0.#}\" y=\"{by + segH / 2:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"middle\" dominant-baseline=\"middle\">{vlabel}</text>");
                            }
                        }
                    }
                    else
                    {
                        // Draw from the zero baseline: positive extends up, negative
                        // extends down. Always emit a non-negative height using the
                        // absolute magnitude (a negative height would clip to zero).
                        var bx = ox + c * groupW + gap + s * pitchW;
                        // Bar spans from the zero baseline (or, on a log axis, the
                        // decade floor) to the value; height is the gap between the
                        // two mapped endpoints. Linear: identical to |val|/span·ph.
                        // Top edge is the smaller Y of the two endpoints; reversed flips
                        // which end that is (with maxMin the baseline is at the TOP so
                        // bars grow downward).
                        var valY = ValToY(val);
                        var bh = Math.Abs(valY - plotZeroY);
                        var by = Math.Min(plotZeroY, valY);
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{bh:0.#}\" {BarFillAttrs(s, c, val)} opacity=\"{FillOpacity(s)}\"/>");
                        if (showDataLabels && !LabelDeleted(s, c))
                        {
                            var vlabel = LabelText(s, c, rawVal, val);
                            // Honor <c:dLblPos> for vertical columns. The value-end tip
                            // is the top edge (by) when the bar grows up, else the bottom
                            // edge (by+bh); the base edge is the opposite. outEnd places
                            // the label just past the tip (Office default); inEnd just
                            // inside the tip; ctr at the bar midpoint; inBase near the
                            // zero baseline. Without this, inEnd and outEnd were identical.
                            var labelAbove = isReversed ? val < 0 : val >= 0;
                            var tipY = labelAbove ? by : by + bh;
                            var baseY = labelAbove ? by + bh : by;
                            double ly;
                            switch (DataLabelPos)
                            {
                                case "inEnd":
                                    ly = labelAbove ? tipY + DataLabelFontPx + 1 : tipY - 3;
                                    break;
                                case "ctr":
                                    ly = by + bh / 2 + DataLabelFontPx / 2.0;
                                    break;
                                case "inBase":
                                    ly = labelAbove ? baseY - 3 : baseY + DataLabelFontPx + 1;
                                    break;
                                default: // outEnd (Office default)
                                    ly = labelAbove ? tipY - 3 : tipY + DataLabelFontPx;
                                    break;
                            }
                            sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{bx + barW / 2:0.#}\" y=\"{ly:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"middle\">{vlabel}</text>");
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
                        var bx = ox + c * groupW + gap + s * pitchW + barW / 2;
                        var byTop = ValToY(rawVal);
                        double plusErr = eb.ValueType == "percentage" ? Math.Abs(rawVal) * eb.Value / 100.0 : errAmount;
                        double minusErr = plusErr;
                        var showPlus = eb.BarType is "both" or "plus";
                        var showMinus = eb.BarType is "both" or "minus";
                        var yTop = showPlus ? ValToY(rawVal + plusErr) : byTop;
                        var yBot = showMinus ? ValToY(rawVal - minusErr) : byTop;
                        sb.AppendLine($"        <line x1=\"{bx:0.#}\" y1=\"{yTop:0.#}\" x2=\"{bx:0.#}\" y2=\"{yBot:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                        if (showPlus && !eb.NoEndCap)
                            sb.AppendLine($"        <line x1=\"{bx - capW:0.#}\" y1=\"{yTop:0.#}\" x2=\"{bx + capW:0.#}\" y2=\"{yTop:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                        if (showMinus && !eb.NoEndCap)
                            sb.AppendLine($"        <line x1=\"{bx - capW:0.#}\" y1=\"{yBot:0.#}\" x2=\"{bx + capW:0.#}\" y2=\"{yBot:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
                    }
                }
            }
            // Trendlines on vertical (column) bar charts. PowerPoint regresses over
            // the 1-based category index and draws the fitted curve across the plot,
            // each category anchored at its group center (ox + (i+0.5)*groupW).
            if (trendlines != null && !stacked)
            {
                for (int s = 0; s < serCount; s++)
                {
                    var tl = s < trendlines.Count ? trendlines[s] : null;
                    if (tl == null) continue;
                    var vals = series[s].values;
                    if (vals.Length < 2) continue;
                    var lineColor = tl.Color ?? colors[s % colors.Count];
                    var xData = new double[vals.Length];
                    var yData = new double[vals.Length];
                    for (int i = 0; i < vals.Length; i++) { xData[i] = i + 1; yData[i] = vals[i]; }
                    Func<double, double> tlMapX = xv => ox + (xv - 0.5) * groupW;
                    AppendTrendline(sb, tl, xData, yData, tlMapX, ValToY, lineColor, ox + pw, oy + 12);
                }
            }
            if (CatAxisVisible)
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var lx = ox + c * groupW + groupW / 2;
                // Vertical columns: category axis is HORIZONTAL at the bottom (y=oy+ph).
                if (TickMarkVisible(CatMajorTickMark))
                    EmitHAxisTick(sb, lx, oy + ph, CatMajorTickMark!);
                if (!CatTickLabelsHidden && (CatTickLabelSkip <= 1 || c % CatTickLabelSkip == 0))
                    EmitBottomAxisLabel(sb, lx, oy + ph + 16, CatColor, catFontSize, label, catLabelRotationDeg, CatTickWeightAttr);
            }
            if (ValAxisVisible)
            for (int t = 0; t <= nTicks; t++)
            {
                var val = isLog ? Math.Pow(logB, logMinExp + t) : niceMin + tickStep * t;
                // BUG1(R25): with an explicit axisMin/max/majorUnit the final
                // tick can land above axisMax (e.g. 450 > 400); real PowerPoint
                // omits any label past the axis top. Skip it.
                if (val > niceMax + 1e-9) continue;
                var label = percentStacked ? $"{(int)val}%" : FmtValAxis(val, valNumFmt);
                // Position the label at the VALUE on the scale (matches the gridlines and
                // bars) rather than an even pixel fraction; no-op when nTicks*tickStep==span.
                var ty = ValToY(val);
                // Vertical columns: value axis is VERTICAL on the left (x=ox).
                if (TickMarkVisible(ValMajorTickMark))
                    EmitVAxisTick(sb, ox, ty, ValMajorTickMark!);
                if (!ValTickLabelsHidden)
                    EmitLeftAxisLabel(sb, ox - 4, ty, AxisColor, valFontSize, label, valLabelRotationDeg, ValTickWeightAttr);
            }
            // Reference-line overlays: vertical bars/columns → horizontal line at value position on the Y (value) axis.
            if (referenceLines != null)
                foreach (var rl in referenceLines)
                {
                    var v = percentStacked ? rl.Value * 100 : rl.Value;
                    if (v < niceMin || v > niceMax) continue;
                    var ry = ValToY(v);
                    var strokeColor = rl.Color.StartsWith("#") ? rl.Color : "#" + rl.Color;
                    var dashArray = RefLineDashArray(rl.Dash);
                    sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{ry:0.#}\" x2=\"{ox + pw}\" y2=\"{ry:0.#}\" stroke=\"{strokeColor}\" stroke-width=\"{rl.WidthPt:0.##}\" stroke-dasharray=\"{dashArray}\"/>");
                }
        }
    }

    // --- Shared decoration primitives (used by both the line and scatter
    // renderers so the two never drift). Each takes pre-computed pixel points
    // and/or value→pixel mappers, so the caller owns axis positioning. ---

    // Catmull-Rom → cubic Bézier smooth path through the given pixel points.
    private static string BuildSmoothPath(IReadOnlyList<(double x, double y)> pts)
    {
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
        return d.ToString();
    }

    // Vertical (Y) error bars at each point. seriesValues feeds stdDev/stdErr.
    private void AppendErrorBars(StringBuilder sb, IReadOnlyList<(double x, double y, double val)> pts,
        ErrorBarInfo eb, double[] seriesValues, Func<double, double> mapY)
    {
        var ebColor = eb.Color ?? "#666";
        var capW = 4.0; // half-width of the cap line

        double errAmount = eb.Value;
        if (eb.ValueType is "stdDev" or "stdErr")
        {
            var mean = seriesValues.Average();
            var variance = seriesValues.Sum(v => (v - mean) * (v - mean)) / seriesValues.Length;
            var stddev = Math.Sqrt(variance);
            errAmount = eb.ValueType == "stdErr" ? stddev / Math.Sqrt(seriesValues.Length) : stddev;
        }

        for (int p = 0; p < pts.Count; p++)
        {
            var val = pts[p].val;
            double plusErr, minusErr;
            if (eb.ValueType == "percentage")
                plusErr = minusErr = Math.Abs(val) * eb.Value / 100.0;
            else
                plusErr = minusErr = errAmount;

            var showPlus = eb.BarType is "both" or "plus";
            var showMinus = eb.BarType is "both" or "minus";

            var yTop = showPlus ? mapY(val + plusErr) : pts[p].y;
            var yBot = showMinus ? mapY(val - minusErr) : pts[p].y;

            sb.AppendLine($"        <line x1=\"{pts[p].x:0.#}\" y1=\"{yTop:0.#}\" x2=\"{pts[p].x:0.#}\" y2=\"{yBot:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
            if (showPlus)
                sb.AppendLine($"        <line x1=\"{pts[p].x - capW:0.#}\" y1=\"{yTop:0.#}\" x2=\"{pts[p].x + capW:0.#}\" y2=\"{yTop:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
            if (showMinus)
                sb.AppendLine($"        <line x1=\"{pts[p].x - capW:0.#}\" y1=\"{yBot:0.#}\" x2=\"{pts[p].x + capW:0.#}\" y2=\"{yBot:0.#}\" stroke=\"{ebColor}\" stroke-width=\"{eb.Width:0.#}\"/>");
        }
    }

    // Regression trendline. xData/yData are the regression domain (category
    // indices for line charts, real X values for scatter); mapXVal maps an
    // xData-domain value to a pixel, mapY maps a Y value to a pixel. Because
    // scatter passes its real X, the fitted slope/equation are correct there
    // (the old line-only path always regressed over the 1-based index).
    private void AppendTrendline(StringBuilder sb, TrendlineInfo tl, double[] xData, double[] yData,
        Func<double, double> mapXVal, Func<double, double> mapY, string lineColor, double fallbackLabelX, double fallbackLabelY)
    {
        if (xData.Length < 2) return;
        // An explicit trendline color (tl.Color, from <c:trendline><c:spPr><a:ln>)
        // arrives as raw OOXML hex with no '#'; emitting stroke="FF0000" is an
        // invalid SVG paint so the curve renders as stroke:none (invisible). The
        // series-color fallback is already '#'-prefixed. CssHexColor is idempotent
        // on '#'-prefixed input, so route both through it. (Affects every chart
        // type's explicit-color trendline, not just area.)
        lineColor = CssHexColor(lineColor);
        var dashArr = tl.Dash != "solid" ? $" stroke-dasharray=\"{RefLineDashArray(tl.Dash)}\"" : "";

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
                var period = Math.Max(2, tl.Period);
                var maPoints = new List<(double x, double y)>();
                for (int i = period - 1; i < xData.Length; i++)
                {
                    double sum = 0;
                    for (int j = 0; j < period; j++) sum += yData[i - j];
                    maPoints.Add((mapXVal(xData[i]), mapY(sum / period)));
                }
                if (maPoints.Count >= 2)
                {
                    var maPath = string.Join(" ", maPoints.Select(p => $"{p.x:0.#},{p.y:0.#}"));
                    sb.AppendLine($"        <polyline points=\"{maPath}\" fill=\"none\" stroke=\"{lineColor}\" stroke-width=\"{tl.Width:0.#}\"{dashArr}/>");
                }
                return; // no equation/R² for moving average
            }
        }

        if (trendFn == null) return;

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
            tlPoints.Add((mapXVal(x), mapY(y)));
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
            var labelX = tlPoints.Count > 0 ? tlPoints[^1].px - 4 : fallbackLabelX;
            var labelY = tlPoints.Count > 0 ? tlPoints[^1].py - 8 : fallbackLabelY;
            sb.AppendLine($"        <text x=\"{labelX:0.#}\" y=\"{labelY:0.#}\" fill=\"{lineColor}\" font-size=\"8\" text-anchor=\"end\" font-style=\"italic\">{HtmlEncode(label)}</text>");
        }
    }

    /// <summary>
    /// Convert raw per-series values into cumulative (stacked) values: series s
    /// at category c becomes the running sum of series 0..s at c. When percent,
    /// each category's stack is first normalized to 100% of the column total.
    /// </summary>
    private static List<(string name, double[] values)> StackSeries(
        List<(string name, double[] values)> series, bool percent)
    {
        if (series.Count == 0) return series;
        int catCount = series.Max(s => s.values.Length);
        var result = new List<(string name, double[] values)>(series.Count);
        for (int s = 0; s < series.Count; s++)
            result.Add((series[s].name, new double[catCount]));
        for (int c = 0; c < catCount; c++)
        {
            double colTotal = 0;
            if (percent)
                for (int s = 0; s < series.Count; s++)
                    colTotal += c < series[s].values.Length ? series[s].values[c] : 0;
            double running = 0;
            for (int s = 0; s < series.Count; s++)
            {
                var v = c < series[s].values.Length ? series[s].values[c] : 0;
                if (percent) v = colTotal > 0 ? v / colTotal * 100.0 : 0;
                running += v;
                result[s].values[c] = running;
            }
        }
        return result;
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
        List<TrendlineInfo?>? trendlines = null, List<ErrorBarInfo?>? errorBars = null,
        bool scatterMarkersOnly = false, bool stacked = false, bool percent = false,
        string? dataLabelNumFmt = null,
        List<string?>? markerFillColors = null, List<string?>? markerLineColors = null,
        bool showSerName = false, bool showCatName = false, bool showVal = true,
        int? catLabelRotationDeg = null, int? valLabelRotationDeg = null,
        List<bool>? lineHide = null)
    {
        bool isLog = logBase.HasValue && logBase.Value > 1;

        // R16-3: stacked / percentStacked line — each series is plotted at the
        // cumulative sum of itself and all series below it. Percent normalizes
        // each category's stack to 100. Transform the series up-front so axis
        // scaling, markers, and labels all reflect the stacked geometry.
        // R44: preserve the pre-stack series so data-label TEXT shows the
        // original per-series value, while geometry/markers/axis use the
        // stacked (cumulative) values. Mirrors the bar path's `rawVal` split:
        // the bar renderer labels the original value (or percentage when
        // showPercent), never the cumulative stack height.
        var originalSeries = series;
        if (stacked || percent)
            series = StackSeries(series, percent);

        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var dataMax = allValues.Max();
        // R12b: dataMin must include negative values. The old `.Where(v => v > 0)`
        // discarded negatives, so a series like [-120,85,-45,210] reported dataMin=85
        // and the axis floor stayed at 0 — negative points then clamped onto the
        // zero baseline. Log scale still needs a positive floor (log of a
        // non-positive value is undefined), so keep the positive-only min there.
        var dataMin = isLog
            ? allValues.Where(v => v > 0).DefaultIfEmpty(1).Min()
            : allValues.Min();
        if (dataMax <= 0 && isLog) dataMax = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        // X position of category c: evenly-spaced slots. (Scatter charts, which
        // value-position X, are handled by RenderScatterChartSvg, not here.)
        double MapX(int c) => ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);

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
            // Min-aware nice axis: when an explicit non-zero axis min is set and
            // no explicit axis max, derive step/top from the VISIBLE range
            // [axisMin, dataMax] so the top doesn't overshoot (e.g. min=50,
            // dataMax=200 → step 20, top 210, not a 0-based 0..300). When
            // axisMin is 0/absent this falls through to the unchanged path.
            if (axisMin.HasValue && axisMin.Value > 0 && !axisMax.HasValue && !majorUnit.HasValue)
                (niceMax, tickStep, nTicks) = ComputeNiceAxisFromMin(axisMin.Value, dataMax);
            else
                (niceMax, tickStep, nTicks) = ComputeNiceAxis(computeMax);
            if (axisMax.HasValue) niceMax = axisMax.Value;
            // R12b: floor the axis at a nice value ≤ 0 when data has negatives,
            // instead of hard-coding 0 (which clamped negative points to the
            // baseline). Mirror the bar-chart axis logic (DataToY path): a
            // negative floor is -ComputeNiceAxis(|dataMin|).niceMax, snapped to
            // the same tickStep so gridlines/labels stay aligned and a tick
            // lands on the negative extreme.
            if (axisMin.HasValue)
                niceMin = axisMin.Value;
            else if (dataMin < 0)
            {
                var negMagnitude = ComputeNiceAxis(-dataMin).niceMax;
                niceMin = -negMagnitude;
            }
            else
                niceMin = 0;
            if (majorUnit.HasValue && majorUnit.Value > 0)
            {
                tickStep = majorUnit.Value;
                // Round the top up to a multiple of majorUnit (headroom above the
                // data), matching PowerPoint — parity with the bar/column path.
                // Skip when an explicit max pins the top or the floor is negative.
                if (!axisMax.HasValue && niceMin >= 0)
                    niceMax = Math.Ceiling(ComputeNiceAxis(dataMax).niceMax / tickStep) * tickStep;
                nTicks = (int)Math.Ceiling((niceMax - niceMin) / tickStep);
            }
            else if (niceMin < 0)
            {
                // Re-derive tick count across the full (negative→positive) span
                // using the existing tickStep so a gridline sits on the zero line
                // and on the negative floor.
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
        if (ShowValGridlines)
        for (int t = 1; t <= nTicks; t++)
        {
            double tickVal = isLog ? niceMin + t : niceMin + tickStep * t;
            if (!isLog && tickVal > niceMax + 1e-9) continue; // no gridline past axisMax
            var gy = MapY(isLog ? Math.Pow(logBase!.Value, tickVal) : tickVal);
            var lineGridDash = !string.IsNullOrEmpty(GridlineDash) && GridlineDash != "solid" ? RefLineDashArray(GridlineDash) : "none";
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"{GridlineWidthPx:0.##}\" stroke-dasharray=\"{lineGridDash}\"/>");
        }
        // Category-axis major gridlines (vertical) — at the category-slot
        // boundaries, only when <c:catAx><c:majorGridlines/> was declared and
        // the category axis is visible (PowerPoint draws none otherwise).
        if (ShowCatGridlines && CatAxisVisible)
        for (int i = 0; i <= catCount; i++)
        {
            var gx = ox + (double)pw * i / Math.Max(catCount, 1);
            sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
        }
        // Category-axis minor gridlines (vertical, at slot boundaries).
        if (ShowCatMinorGridlines && CatAxisVisible)
        for (int i = 0; i <= catCount; i++)
        {
            var gx = ox + (double)pw * i / Math.Max(catCount, 1);
            sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.25\" opacity=\"0.5\"/>");
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
                var px = MapX(c);
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

            // R16c: scatterStyle=marker draws dots only — skip the connecting
            // line/path entirely. Markers are still emitted below. Also skip when this
            // SERIES has a:ln/a:noFill (PowerPoint "No line" — markers only for that series).
            if (scatterMarkersOnly || (lineHide != null && s < lineHide.Count && lineHide[s]))
            {
                // no line
            }
            else if (isSmooth && pts.Count >= 2)
            {
                var d = BuildSmoothPath(pts.Select(p => (p.x, p.y)).ToList());
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
            var mFill = markerFillColors != null && s < markerFillColors.Count ? markerFillColors[s] : null;
            var mStroke = markerLineColors != null && s < markerLineColors.Count ? markerLineColors[s] : null;
            for (int p = 0; p < pts.Count; p++)
            {
                sb.AppendLine($"        {RenderMarkerSvg(shape, pts[p].x, pts[p].y, mSize, mFill ?? lineColor, mStroke ?? lineColor)}");
                if (showDataLabels && !LabelDeleted(s, p))
                {
                    // R44: label TEXT uses the original (pre-stack) per-series
                    // value; the label POSITION stays at the stacked vertex
                    // (pts[p]). For non-stacked charts originalSeries == series,
                    // so pts[p].val and the original value are identical.
                    var val = (stacked || percent) && s < originalSeries.Count
                              && p < originalSeries[s].values.Length
                        ? originalSeries[s].values[p]
                        : pts[p].val;
                    // BUG5(R25): honor <c:dLbls><c:numFmt> on the data labels
                    // (e.g. "$#,##0"); fall back to the value-axis numFmt (mirrors
                    // the bar LabelText path) then the bare-integer shortcut.
                    var valuePart = !string.IsNullOrEmpty(dataLabelNumFmt) ? FormatAxisValue(val, dataLabelNumFmt)
                        : !string.IsNullOrEmpty(valNumFmt) ? FormatAxisValue(val, valNumFmt)
                        : val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                    // Compose enabled label parts (series name, category name,
                    // value) in PowerPoint's order. Default (showVal only) keeps
                    // the legacy value-only label.
                    var lparts = new List<string>();
                    if (showSerName && s < series.Count) lparts.Add(series[s].name);
                    if (showCatName && p >= 0 && p < categories.Length) lparts.Add(categories[p]);
                    if (showVal) lparts.Add(valuePart);
                    var vlabel = string.Join(", ", lparts);
                    // Honor <c:dLblPos> for line markers (ctr|l|r|t|b). Default is above
                    // the point (PowerPoint's line default); previously the y-6 above
                    // offset was hardcoded, so an explicit dLblPos=b/l/r/ctr was ignored
                    // (bar/pie already honor DataLabelPos).
                    double lblX = pts[p].x, lblY = pts[p].y - 6;
                    var lblAnchor = "middle";
                    if (HasExplicitDataLabelPos)
                    {
                        switch (DataLabelPos)
                        {
                            case "b": lblY = pts[p].y + DataLabelFontPx + 4; break;
                            case "ctr": lblY = pts[p].y + DataLabelFontPx / 3.0; break;
                            case "l": lblX = pts[p].x - 6; lblY = pts[p].y + DataLabelFontPx / 3.0; lblAnchor = "end"; break;
                            case "r": lblX = pts[p].x + 6; lblY = pts[p].y + DataLabelFontPx / 3.0; lblAnchor = "start"; break;
                            default: break; // t / above (legacy)
                        }
                    }
                    sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{lblX:0.#}\" y=\"{lblY:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"{lblAnchor}\">{vlabel}</text>");
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
                AppendErrorBars(sb, allPoints[s], eb, series[s].values, MapY);
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
                // Line/category charts regress over the 1-based category index
                // (xData[i]=i+1); tlMapX converts that index back to a pixel.
                var xData = new double[pts.Count];
                var yData = new double[pts.Count];
                for (int i = 0; i < pts.Count; i++) { xData[i] = i + 1; yData[i] = series[s].values[i]; }
                Func<double, double> tlMapX = xv => ox + (catCount > 1 ? pw * (xv - 1) / (catCount - 1) : pw / 2.0);
                AppendTrendline(sb, tl, xData, yData, tlMapX, MapY, lineColor, ox + pw, oy + 12);
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

        // Category labels (+ major tick marks below the bottom axis)
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            if (CatAxisVisible && TickMarkVisible(CatMajorTickMark))
                EmitHAxisTick(sb, MapX(c), oy + ph, CatMajorTickMark!);
            // Honor <c:catAx><c:txPr><a:bodyPr rot> via EmitBottomAxisLabel (mirrors
            // the bar renderer). Previously the line/area renderers emitted a raw
            // horizontal <text>, dropping the category-axis label rotation that the
            // bar chart already applied (the bottom-margin reservation already fired
            // for all chart types, leaving the labels un-rotated in the gap).
            if (!CatTickLabelsHidden && (CatTickLabelSkip <= 1 || c % CatTickLabelSkip == 0))
                EmitBottomAxisLabel(sb, MapX(c), oy + ph + 16, CatColor, CatFontPx, label, catLabelRotationDeg, CatTickWeightAttr);
        }

        // Value axis labels (+ major tick marks left of the value axis)
        for (int t = 0; t <= nTicks; t++)
        {
            double tickVal;
            string label;
            if (isLog)
            {
                var exp = niceMin + t;
                tickVal = Math.Pow(logBase!.Value, exp);
                label = FmtValAxis(tickVal, valNumFmt);
            }
            else
            {
                tickVal = niceMin + tickStep * t;
                if (tickVal > niceMax + 1e-9) continue; // no label past axisMax
                label = FmtValAxis(tickVal, valNumFmt);
            }
            var ty = MapY(tickVal);
            if (ValAxisVisible && TickMarkVisible(ValMajorTickMark))
                EmitVAxisTick(sb, ox, ty, ValMajorTickMark!);
            if (!ValTickLabelsHidden)
                EmitLeftAxisLabel(sb, ox - 4, ty, AxisColor, ValFontPx, label, valLabelRotationDeg, ValTickWeightAttr);
        }
    }

    public void RenderPieChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH, double holeRatio = 0.0, bool showDataLabels = false,
        bool showVal = false, bool showPercent = false, bool showCatName = false, List<double>? explosions = null,
        bool showSerName = false)
    {
        var values = series.FirstOrDefault().values ?? [];
        if (values.Length == 0) return;
        var total = values.Sum();
        if (total <= 0) return;

        // PowerPoint draws a thin separator (the chart-area background, white by
        // default) between adjacent pie/doughnut slices. Without it slices touch
        // edge-to-edge — invisible for monochrome (varyColors=false) pies. Emit a
        // default white slice stroke to match.
        const string sliceStroke = " stroke=\"#FFFFFF\" stroke-width=\"2\"";

        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var r = Math.Min(svgW, svgH) * 0.42;
        // Slice explosion: each slice's center shifts outward along its bisector
        // by explosion% of the radius. Shrink the base radius so the most-exploded
        // slice still fits inside the plot area (PowerPoint scales the pie down).
        var maxExpl = explosions != null && explosions.Count > 0 ? explosions.Max() : 0.0;
        if (maxExpl > 0) r /= (1 + maxExpl);
        var innerR = r * holeRatio;

        // Multi-series DOUGHNUT: PowerPoint draws one concentric ring per series,
        // sharing the center hole, each ring an equal-width band of the annulus
        // [innerR, r]. Series 0 is OUTERMOST. Each ring colors its slices by
        // category index (the same per-category palette as a single ring).
        // Single-series doughnut and all pie charts (holeRatio == 0) fall through
        // to the original single-ring path below, unchanged.
        if (holeRatio > 0 && series.Count > 1)
        {
            RenderMultiRingDoughnut(sb, series, categories, colors, cx, cy, r, innerR,
                showDataLabels, showVal, showPercent, showCatName);
            return;
        }
        // firstSliceAng rotates the start edge clockwise from 12 o'clock. SVG y
        // grows downward, so a clockwise rotation adds to the angle directly.
        var firstSliceOffset = FirstSliceAngle * Math.PI / 180.0;
        var startAngle = -Math.PI / 2 + firstSliceOffset;

        for (int i = 0; i < values.Length; i++)
        {
            var sliceAngle = 2 * Math.PI * values[i] / total;
            var endAngle = startAngle + sliceAngle;
            var color = i < colors.Count ? colors[i] : DefaultColors[i % DefaultColors.Length];

            // Per-slice exploded center: push out along the slice bisector.
            var expl = explosions != null && i < explosions.Count ? explosions[i] : 0.0;
            var midAngle = startAngle + sliceAngle / 2;
            var cx0 = expl > 0 ? cx + r * expl * Math.Cos(midAngle) : cx;
            var cy0 = expl > 0 ? cy + r * expl * Math.Sin(midAngle) : cy;

            var sliceOpacity = FillOpacity(i);
            if (values.Length == 1 && holeRatio <= 0)
                sb.AppendLine($"        <circle cx=\"{cx0:0.#}\" cy=\"{cy0:0.#}\" r=\"{r:0.#}\" fill=\"{color}\" opacity=\"{sliceOpacity}\"/>");
            else if (holeRatio > 0)
            {
                // A single ring segment spanning ~full circle (one data point,
                // or one value ≈ total) has start≈end, so a single SVG arc is
                // zero-length and browsers draw nothing. Split into two full-ring
                // halves drawn with the even-odd fill rule (outer circle minus
                // inner circle) so the annulus renders. Threshold a hair under 2π
                // to catch float rounding.
                if (sliceAngle >= 2 * Math.PI - 1e-6)
                {
                    sb.AppendLine($"        <path d=\"M {cx0 - r:0.#},{cy0:0.#} A {r:0.#},{r:0.#} 0 1,1 {cx0 + r:0.#},{cy0:0.#} A {r:0.#},{r:0.#} 0 1,1 {cx0 - r:0.#},{cy0:0.#} Z M {cx0 - innerR:0.#},{cy0:0.#} A {innerR:0.#},{innerR:0.#} 0 1,1 {cx0 + innerR:0.#},{cy0:0.#} A {innerR:0.#},{innerR:0.#} 0 1,1 {cx0 - innerR:0.#},{cy0:0.#} Z\" fill=\"{color}\" fill-rule=\"evenodd\" opacity=\"{sliceOpacity}\"{sliceStroke}/>");
                }
                else
                {
                    var ox1 = cx0 + r * Math.Cos(startAngle); var oy1 = cy0 + r * Math.Sin(startAngle);
                    var ox2 = cx0 + r * Math.Cos(endAngle); var oy2 = cy0 + r * Math.Sin(endAngle);
                    var ix1 = cx0 + innerR * Math.Cos(endAngle); var iy1 = cy0 + innerR * Math.Sin(endAngle);
                    var ix2 = cx0 + innerR * Math.Cos(startAngle); var iy2 = cy0 + innerR * Math.Sin(startAngle);
                    var largeArc = sliceAngle > Math.PI ? 1 : 0;
                    sb.AppendLine($"        <path d=\"M {ox1:0.#},{oy1:0.#} A {r:0.#},{r:0.#} 0 {largeArc},1 {ox2:0.#},{oy2:0.#} L {ix1:0.#},{iy1:0.#} A {innerR:0.#},{innerR:0.#} 0 {largeArc},0 {ix2:0.#},{iy2:0.#} Z\" fill=\"{color}\" opacity=\"{sliceOpacity}\"{sliceStroke}/>");
                }
            }
            else
            {
                var x1 = cx0 + r * Math.Cos(startAngle); var y1 = cy0 + r * Math.Sin(startAngle);
                var x2 = cx0 + r * Math.Cos(endAngle); var y2 = cy0 + r * Math.Sin(endAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {cx0:0.#},{cy0:0.#} L {x1:0.#},{y1:0.#} A {r:0.#},{r:0.#} 0 {largeArc},1 {x2:0.#},{y2:0.#} Z\" fill=\"{color}\" opacity=\"{sliceOpacity}\"{sliceStroke}/>");
            }
            startAngle = endAngle;
        }
        if (showDataLabels)
        {
            var labelAngle = -Math.PI / 2 + firstSliceOffset;
            // <c:dLblPos val="outEnd"> places labels just beyond the pie edge along
            // each slice bisector; inEnd/ctr/bestFit keep them inside the slice.
            // OOXML's default pie/doughnut label position is bestFit (ON the
            // segment), NOT outEnd. The "outEnd" default is bar/column-appropriate,
            // so only treat outEnd as "outside" for pie/doughnut when the XML
            // explicitly declared <c:dLblPos val="outEnd"/>. This matches PowerPoint,
            // which draws default labels on the colored slices (readable even on a
            // dark chart background) rather than outside in dark text.
            var labelOutside = HasExplicitDataLabelPos && DataLabelPos == "outEnd";
            var labelR = labelOutside ? r * 1.12
                : holeRatio > 0 ? r * (1 + holeRatio) / 2 : r * 0.65;
            for (int i = 0; i < values.Length; i++)
            {
                var sliceAngle = 2 * Math.PI * values[i] / total;
                var midAngle = labelAngle + sliceAngle / 2;
                var expl = explosions != null && i < explosions.Count ? explosions[i] : 0.0;
                var lcx = expl > 0 ? cx + r * expl * Math.Cos(midAngle) : cx;
                var lcy = expl > 0 ? cy + r * expl * Math.Sin(midAngle) : cy;
                var lx = lcx + labelR * Math.Cos(midAngle);
                var ly = lcy + labelR * Math.Sin(midAngle);
                var pct = values[i] / total * 100;
                string label;
                if (showVal && !showPercent)
                    label = pct >= 5 ? $"{values[i]:0.##}" : "";
                else if (showPercent && !showVal)
                    label = pct >= 5 ? $"{pct:0}%" : "";
                else if (showVal && showPercent)
                    label = pct >= 5 ? $"{values[i]:0.##} ({pct:0}%)" : "";
                else if ((showCatName || showSerName) && !showVal && !showPercent)
                    label = ""; // name-only — value text intentionally empty (name parts prepended below)
                else
                    label = pct >= 5 ? $"{pct:0}%" : ""; // default to percent for pie
                // showCatName / showSerName: prepend the category and/or series
                // name (PowerPoint style: "Series, Category, value, pct").
                // Honored independently of val/percent.
                if (showCatName && pct >= 5 && i < categories.Length && !string.IsNullOrEmpty(categories[i]))
                    label = string.IsNullOrEmpty(label) ? categories[i] : $"{categories[i]}, {label}";
                if (showSerName && pct >= 5 && series.Count > 0 && !string.IsNullOrEmpty(series[0].name))
                    label = string.IsNullOrEmpty(label) ? series[0].name : $"{series[0].name}, {label}";
                // Outside labels sit on the plot background, not on a colored
                // slice — use a dark fill (white is invisible there).
                var labelFill = labelOutside ? "#444" : "#fff";
                if (!string.IsNullOrEmpty(label))
                    sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"{labelFill}\" font-size=\"{DataLabelFontPx}\" font-weight=\"bold\" text-anchor=\"middle\" dominant-baseline=\"central\">{label}</text>");
                labelAngle += sliceAngle;
            }
        }
    }

    /// <summary>
    /// Multi-series doughnut: one concentric ring per series. The annulus
    /// [innerR, r] is split into N equal-width bands; series[0] occupies the
    /// OUTERMOST band, series[N-1] the innermost (nearest the hole), matching
    /// PowerPoint. Each ring colors its slices by category index using the
    /// shared per-category palette. Data labels are drawn on the outer ring
    /// only (PowerPoint labels every ring, but the outer ring keeps single-ring
    /// label behavior intact and avoids label crowding in the thin inner bands).
    /// </summary>
    private void RenderMultiRingDoughnut(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, double cx, double cy, double r, double innerR,
        bool showDataLabels, bool showVal, bool showPercent, bool showCatName)
    {
        var firstSliceOffset = FirstSliceAngle * Math.PI / 180.0;
        var bandWidth = (r - innerR) / series.Count;
        // Default white separator between adjacent slices (see RenderPieChartSvg).
        const string sliceStroke = " stroke=\"#FFFFFF\" stroke-width=\"2\"";

        for (int s = 0; s < series.Count; s++)
        {
            var values = series[s].values ?? [];
            if (values.Length == 0) continue;
            var total = values.Sum();
            if (total <= 0) continue;

            // Series 0 = outermost band. Band s spans [bandInner, bandOuter].
            var bandOuter = r - s * bandWidth;
            var bandInner = bandOuter - bandWidth;
            var startAngle = -Math.PI / 2 + firstSliceOffset;

            for (int i = 0; i < values.Length; i++)
            {
                var sliceAngle = 2 * Math.PI * values[i] / total;
                var endAngle = startAngle + sliceAngle;
                var color = i < colors.Count ? colors[i] : DefaultColors[i % DefaultColors.Length];
                var sliceOpacity = FillOpacity(i);

                if (sliceAngle >= 2 * Math.PI - 1e-6)
                {
                    sb.AppendLine($"        <path d=\"M {cx - bandOuter:0.#},{cy:0.#} A {bandOuter:0.#},{bandOuter:0.#} 0 1,1 {cx + bandOuter:0.#},{cy:0.#} A {bandOuter:0.#},{bandOuter:0.#} 0 1,1 {cx - bandOuter:0.#},{cy:0.#} Z M {cx - bandInner:0.#},{cy:0.#} A {bandInner:0.#},{bandInner:0.#} 0 1,1 {cx + bandInner:0.#},{cy:0.#} A {bandInner:0.#},{bandInner:0.#} 0 1,1 {cx - bandInner:0.#},{cy:0.#} Z\" fill=\"{color}\" fill-rule=\"evenodd\" opacity=\"{sliceOpacity}\"{sliceStroke}/>");
                }
                else
                {
                    var ox1 = cx + bandOuter * Math.Cos(startAngle); var oy1 = cy + bandOuter * Math.Sin(startAngle);
                    var ox2 = cx + bandOuter * Math.Cos(endAngle); var oy2 = cy + bandOuter * Math.Sin(endAngle);
                    var ix1 = cx + bandInner * Math.Cos(endAngle); var iy1 = cy + bandInner * Math.Sin(endAngle);
                    var ix2 = cx + bandInner * Math.Cos(startAngle); var iy2 = cy + bandInner * Math.Sin(startAngle);
                    var largeArc = sliceAngle > Math.PI ? 1 : 0;
                    sb.AppendLine($"        <path d=\"M {ox1:0.#},{oy1:0.#} A {bandOuter:0.#},{bandOuter:0.#} 0 {largeArc},1 {ox2:0.#},{oy2:0.#} L {ix1:0.#},{iy1:0.#} A {bandInner:0.#},{bandInner:0.#} 0 {largeArc},0 {ix2:0.#},{iy2:0.#} Z\" fill=\"{color}\" opacity=\"{sliceOpacity}\"{sliceStroke}/>");
                }
                startAngle = endAngle;
            }
        }

        // Labels on the outermost ring only (series 0).
        if (showDataLabels)
        {
            var outerVals = series[0].values ?? [];
            var outerTotal = outerVals.Sum();
            if (outerTotal <= 0) return;
            var labelR = r - bandWidth / 2;
            var labelAngle = -Math.PI / 2 + firstSliceOffset;
            for (int i = 0; i < outerVals.Length; i++)
            {
                var sliceAngle = 2 * Math.PI * outerVals[i] / outerTotal;
                var midAngle = labelAngle + sliceAngle / 2;
                var lx = cx + labelR * Math.Cos(midAngle);
                var ly = cy + labelR * Math.Sin(midAngle);
                var pct = outerVals[i] / outerTotal * 100;
                string label;
                if (showVal && !showPercent) label = pct >= 5 ? $"{outerVals[i]:0.##}" : "";
                else if (showPercent && !showVal) label = pct >= 5 ? $"{pct:0}%" : "";
                else if (showVal && showPercent) label = pct >= 5 ? $"{outerVals[i]:0.##} ({pct:0}%)" : "";
                else if (showCatName && !showVal && !showPercent) label = "";
                else label = pct >= 5 ? $"{pct:0}%" : "";
                if (showCatName && pct >= 5 && i < categories.Length && !string.IsNullOrEmpty(categories[i]))
                    label = string.IsNullOrEmpty(label) ? categories[i] : $"{categories[i]}, {label}";
                if (!string.IsNullOrEmpty(label))
                    sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"#fff\" font-size=\"{DataLabelFontPx}\" font-weight=\"bold\" text-anchor=\"middle\" dominant-baseline=\"central\">{label}</text>");
                labelAngle += sliceAngle;
            }
        }
    }

    /// <summary>
    /// pieOfPie / barOfPie composite. The single data series is split into a
    /// "main" group (the leading points) and a "secondary" group (the trailing
    /// points, default the last 3 as PowerPoint does). The main pie shows the
    /// leading slices plus one aggregate slice for the secondary group; the
    /// secondary group is rendered as its own small pie (pieOfPie) or a vertical
    /// bar stack (barOfPie), joined to the aggregate slice by connector lines.
    /// </summary>
    public void RenderOfPieChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool isBar,
        bool showDataLabels = false, bool showVal = true, bool showPercent = false,
        string? dataLabelNumFmt = null)
    {
        var values = series.FirstOrDefault().values ?? [];
        if (values.Length == 0) return;
        var total = values.Sum();
        if (total <= 0) return;

        // Format a single slice value as a data label, honoring showVal/showPercent
        // and the dLbls numFmt (mirrors the standard pie renderer's label content).
        string SliceLabel(double v, double tot)
        {
            var pct = tot > 0 ? v / tot * 100 : 0;
            var valTxt = !string.IsNullOrEmpty(dataLabelNumFmt) ? FormatAxisValue(v, dataLabelNumFmt) : $"{v:0.##}";
            if (showVal && showPercent) return $"{valTxt} ({pct:0}%)";
            if (showPercent) return $"{pct:0}%";
            return valTxt;
        }

        // Split: trailing `secCount` points go to the secondary plot.
        int secCount = Math.Min(3, Math.Max(1, values.Length - 1));
        int mainCount = values.Length - secCount;
        var mainVals = values.Take(mainCount).ToList();
        var secVals = values.Skip(mainCount).ToList();
        var secSum = secVals.Sum();

        // ── Main pie (left half of the plot) — leading slices + aggregate ──
        var mcx = ox + pw * 0.30;
        var mcy = oy + ph / 2.0;
        var mr = Math.Min(pw * 0.30, ph * 0.42);
        var startAngle = -Math.PI / 2;
        var aggMidAngle = startAngle; // updated when we draw the aggregate slice

        // Slices = leading points, then one aggregate slice for the secondary group.
        var mainSlices = new List<double>(mainVals) { secSum };
        for (int i = 0; i < mainSlices.Count; i++)
        {
            var sliceAngle = 2 * Math.PI * mainSlices[i] / total;
            var endAngle = startAngle + sliceAngle;
            var color = i < colors.Count ? colors[i] : DefaultColors[i % DefaultColors.Length];
            var x1 = mcx + mr * Math.Cos(startAngle); var y1 = mcy + mr * Math.Sin(startAngle);
            var x2 = mcx + mr * Math.Cos(endAngle); var y2 = mcy + mr * Math.Sin(endAngle);
            var largeArc = sliceAngle > Math.PI ? 1 : 0;
            sb.AppendLine($"        <path d=\"M {mcx:0.#},{mcy:0.#} L {x1:0.#},{y1:0.#} A {mr:0.#},{mr:0.#} 0 {largeArc},1 {x2:0.#},{y2:0.#} Z\" fill=\"{color}\" opacity=\"{FillOpacity(i)}\"/>");
            if (showDataLabels)
            {
                var midAngle = startAngle + sliceAngle / 2;
                var lx = mcx + mr * 0.65 * Math.Cos(midAngle);
                var ly = mcy + mr * 0.65 * Math.Sin(midAngle);
                sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"#fff\" font-size=\"{DataLabelFontPx}\" font-weight=\"bold\" text-anchor=\"middle\" dominant-baseline=\"central\">{HtmlEncode(SliceLabel(mainSlices[i], total))}</text>");
            }
            if (i == mainSlices.Count - 1) aggMidAngle = startAngle + sliceAngle / 2;
            startAngle = endAngle;
        }

        // Connector lines from the aggregate slice edge to the secondary plot.
        var aggEdgeX = mcx + mr * Math.Cos(aggMidAngle);
        var aggEdgeY = mcy + mr * Math.Sin(aggMidAngle);
        var secX = ox + pw * 0.72;
        var secTop = oy + ph * 0.20;
        var secBot = oy + ph * 0.80;
        sb.AppendLine($"        <line x1=\"{aggEdgeX:0.#}\" y1=\"{aggEdgeY:0.#}\" x2=\"{secX:0.#}\" y2=\"{secTop:0.#}\" stroke=\"#999\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{aggEdgeX:0.#}\" y1=\"{aggEdgeY:0.#}\" x2=\"{secX:0.#}\" y2=\"{secBot:0.#}\" stroke=\"#999\" stroke-width=\"1\"/>");

        if (secSum <= 0) return;

        if (isBar)
        {
            // ── Secondary bar stack ──
            var barW = pw * 0.12;
            var barX = secX;
            var stackH = secBot - secTop;
            var yCursor = secBot;
            for (int i = 0; i < secVals.Count; i++)
            {
                var h = stackH * secVals[i] / secSum;
                var color = (mainCount + i) < colors.Count ? colors[mainCount + i]
                    : DefaultColors[(mainCount + i) % DefaultColors.Length];
                sb.AppendLine($"        <rect x=\"{barX:0.#}\" y=\"{yCursor - h:0.#}\" width=\"{barW:0.#}\" height=\"{h:0.#}\" fill=\"{color}\" opacity=\"{FillOpacity(mainCount + i)}\"/>");
                if (showDataLabels)
                    sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{barX + barW / 2:0.#}\" y=\"{yCursor - h / 2:0.#}\" fill=\"#fff\" font-size=\"{DataLabelFontPx}\" font-weight=\"bold\" text-anchor=\"middle\" dominant-baseline=\"central\">{HtmlEncode(SliceLabel(secVals[i], secSum))}</text>");
                yCursor -= h;
            }
        }
        else
        {
            // ── Secondary small pie ──
            var scx = secX;
            var scy = oy + ph / 2.0;
            var sr = Math.Min(pw * 0.18, ph * 0.28);
            var sStart = -Math.PI / 2;
            for (int i = 0; i < secVals.Count; i++)
            {
                var sliceAngle = 2 * Math.PI * secVals[i] / secSum;
                var endAngle = sStart + sliceAngle;
                var color = (mainCount + i) < colors.Count ? colors[mainCount + i]
                    : DefaultColors[(mainCount + i) % DefaultColors.Length];
                var x1 = scx + sr * Math.Cos(sStart); var y1 = scy + sr * Math.Sin(sStart);
                var x2 = scx + sr * Math.Cos(endAngle); var y2 = scy + sr * Math.Sin(endAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {scx:0.#},{scy:0.#} L {x1:0.#},{y1:0.#} A {sr:0.#},{sr:0.#} 0 {largeArc},1 {x2:0.#},{y2:0.#} Z\" fill=\"{color}\" opacity=\"{FillOpacity(mainCount + i)}\"/>");
                if (showDataLabels)
                {
                    var midAngle = sStart + sliceAngle / 2;
                    var lx = scx + sr * 0.65 * Math.Cos(midAngle);
                    var ly = scy + sr * 0.65 * Math.Sin(midAngle);
                    sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"#fff\" font-size=\"{DataLabelFontPx}\" font-weight=\"bold\" text-anchor=\"middle\" dominant-baseline=\"central\">{HtmlEncode(SliceLabel(secVals[i], secSum))}</text>");
                }
                sStart = endAngle;
            }
        }
    }

    public void RenderAreaChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool stacked = false,
        bool percent = false,
        double? axisMin = null, double? axisMax = null, double? majorUnit = null, string? valNumFmt = null,
        bool showDataLabels = false, bool showVal = true, bool showSerName = false,
        bool showCatName = false, string? dataLabelNumFmt = null,
        int? catLabelRotationDeg = null, int? valLabelRotationDeg = null,
        bool isReversed = false, List<TrendlineInfo?>? trendlines = null)
    {
        if (series.Count == 0) return;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        if (catCount == 0) return;

        // R16-2: percentStacked normalizes each category's stack to 100% of the
        // column total; the axis is fixed at 0..100. Percent implies stacked.
        if (percent) stacked = true;

        var cumulative = new double[series.Count, catCount];
        for (int c = 0; c < catCount; c++)
        {
            double colTotal = 0;
            if (percent)
                for (int s = 0; s < series.Count; s++)
                    colTotal += c < series[s].values.Length ? series[s].values[c] : 0;
            double runningSum = 0;
            for (int s = 0; s < series.Count; s++)
            {
                var val = c < series[s].values.Length ? series[s].values[c] : 0;
                if (percent) val = colTotal > 0 ? val / colTotal * 100.0 : 0;
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
        var (niceMax, tickInterval, tickCount) = percent
            ? (100.0, 20.0, 5)
            : ComputeNiceAxis(Math.Abs(maxVal) > Math.Abs(minVal) ? maxVal : -minVal);
        // For non-stacked charts with negative values, expand the axis to cover minVal.
        // niceMin straddles zero so the zero line sits inside the plot and negative
        // areas fill below it — same domain rule as the bar/column path
        // (R12b parity: dataMin = Math.Min(0, allValues.Min())).
        var niceMin = minVal < 0 ? -ComputeNiceAxis(-minVal).niceMax : 0.0;
        // BUG2(R25): honor explicit axis scaling (axisMin/axisMax/majorUnit) like
        // the column renderer, so an area chart with set axisMax=… isn't ignored.
        if (!percent)
        {
            if (axisMax.HasValue) niceMax = axisMax.Value;
            if (axisMin.HasValue) niceMin = axisMin.Value;
            if (majorUnit.HasValue && majorUnit.Value > 0)
            {
                tickInterval = majorUnit.Value;
                tickCount = (int)Math.Ceiling((niceMax - niceMin) / tickInterval);
            }
        }
        var axisRange = niceMax - niceMin;
        // Ticks/gridlines must span the whole niceMin..niceMax range, not just the
        // positive 0..niceMax portion. nTicks counts steps across the full domain so
        // labels read e.g. -4,-2,0,2,4,6 instead of 0..6 with negatives clipped.
        var nTicks = percent ? tickCount : (int)Math.Round((niceMax - niceMin) / tickInterval);
        if (nTicks < 1) nTicks = 1;

        // Helper: map a data value to a y-coordinate within [oy, oy+ph]. When the value
        // axis is reversed (<c:scaling><c:orientation val="maxMin"/>), the mapping flips
        // so max sits at the bottom — mirrors RenderLineChartSvg's isReversed MapY.
        // Previously area ignored isReversed (no param) and rendered upside-down vs
        // PowerPoint. Gridlines and value-axis tick labels route through DataToY so they
        // flip consistently. (Category labels stay at the bottom, matching the line/bar
        // renderers' reversed behavior.)
        double DataToY(double v) => isReversed
            ? oy + (v - niceMin) / axisRange * ph
            : oy + ph - (v - niceMin) / axisRange * ph;
        double ZeroY() => DataToY(0.0);
        // Like DataToY but clamped to the plot rect — for stacked-area polygons whose
        // baseline (data value 0) can fall below an explicit axisMin (PowerPoint clips
        // the fill to the axis bottom rather than letting it run off-plot).
        double ClampY(double v) => Math.Clamp(DataToY(v), oy, oy + ph);

        // Gridlines at tick VALUES (niceMin + tickInterval*t), matching the value-axis
        // LABELS below (which use the same expression) and the area fills. Previously the
        // gridlines stepped by axisRange/nTicks, which diverges from tickInterval*t when an
        // explicit axisMax isn't a multiple of tickInterval. No-op when they coincide.
        if (ShowValGridlines && ValAxisVisible)
        for (int t = 1; t <= nTicks; t++)
        {
            var gridVal = niceMin + tickInterval * t;
            if (gridVal > niceMax + 1e-9) continue;
            var gy = DataToY(gridVal);
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"{GridlineWidthPx:0.##}\"{ValGridDashAttr}/>");
        }
        if (ShowValMinorGridlines && ValAxisVisible)
        for (int t = 0; t < nTicks; t++)
            for (int m = 1; m < MinorGridlineCount; m++)
            {
                var minorVal = niceMin + tickInterval * (t + (double)m / MinorGridlineCount);
                if (minorVal > niceMax + 1e-9) continue;
                var gy = DataToY(minorVal);
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.25\" opacity=\"0.5\"/>");
            }
        // Category-axis major gridlines (vertical) — at the category-slot
        // boundaries, gated on <c:catAx><c:majorGridlines/> + category-axis
        // visibility (PowerPoint draws none by default).
        if (ShowCatGridlines && CatAxisVisible)
        for (int i = 0; i <= catCount; i++)
        {
            var gx = ox + (double)pw * i / Math.Max(catCount, 1);
            sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
        }
        // Category-axis minor gridlines (vertical, at slot boundaries).
        if (ShowCatMinorGridlines && CatAxisVisible)
        for (int i = 0; i <= catCount; i++)
        {
            var gx = ox + (double)pw * i / Math.Max(catCount, 1);
            sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.25\" opacity=\"0.5\"/>");
        }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        // Zero baseline when the domain straddles zero (negative data present) — the
        // area fills meet at this line, positives above / negatives below.
        if (niceMin < 0)
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{ZeroY():0.#}\" x2=\"{ox + pw}\" y2=\"{ZeroY():0.#}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        if (stacked)
        {
            for (int s = series.Count - 1; s >= 0; s--)
            {
                var topPoints = new List<string>();
                var bottomPoints = new List<string>();
                for (int c = 0; c < catCount; c++)
                {
                    var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                    topPoints.Add($"{px:0.#},{ClampY(cumulative[s, c]):0.#}");
                    var bottomVal = s > 0 ? cumulative[s - 1, c] : 0;
                    bottomPoints.Add($"{px:0.#},{ClampY(bottomVal):0.#}");
                }
                bottomPoints.Reverse();
                sb.AppendLine($"        <polygon points=\"{string.Join(" ", topPoints)} {string.Join(" ", bottomPoints)}\" fill=\"{colors[s % colors.Count]}\" opacity=\"{FillOpacity(s)}\"/>");
            }
        }
        else
        {
            var baseY = ZeroY();
            // Forward index order: series0 bottom-most, seriesN on top — matches native PowerPoint
            var renderOrder = Enumerable.Range(0, series.Count).ToList();
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
                sb.AppendLine($"        <polygon points=\"{firstX:0.#},{baseY:0.#} {string.Join(" ", topPoints)} {lastX:0.#},{baseY:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"{FillOpacity(s)}\"/>");
            }
        }
        if (CatAxisVisible)
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            if (TickMarkVisible(CatMajorTickMark))
                EmitHAxisTick(sb, lx, oy + ph, CatMajorTickMark!);
            // Honor the category-axis label rotation (mirrors bar/line via EmitBottomAxisLabel).
            if (!CatTickLabelsHidden && (CatTickLabelSkip <= 1 || c % CatTickLabelSkip == 0))
                EmitBottomAxisLabel(sb, lx, oy + ph + 16, CatColor, CatFontPx, label, catLabelRotationDeg, CatTickWeightAttr);
        }
        if (ValAxisVisible)
        for (int t = 0; t <= nTicks; t++)
        {
            // BUG2(R25): start at niceMin and format with the value-axis numFmt
            // (e.g. "$#,##0") instead of hardcoded integer text.
            var val = niceMin + tickInterval * t;
            if (val > niceMax + 1e-9) continue; // BUG1(R25): no label past axisMax
            var label = percent ? (val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}")
                : FmtValAxis(val, valNumFmt);
            var ty = DataToY(val);
            if (TickMarkVisible(ValMajorTickMark))
                EmitVAxisTick(sb, ox, ty, ValMajorTickMark!);
            if (!ValTickLabelsHidden)
                EmitLeftAxisLabel(sb, ox - 4, ty, AxisColor, ValFontPx, label, valLabelRotationDeg, ValTickWeightAttr);
        }

        // Data labels at each vertex (parity with bar/line/pie). The label TEXT is
        // the original per-series value; the label POSITION sits at the plotted
        // vertex — for stacked/percent that is the cumulative top, for non-stacked
        // the raw value. Emitted last so labels sit above fills and gridlines.
        if (showDataLabels)
        for (int s = 0; s < series.Count; s++)
            for (int c = 0; c < catCount; c++)
            {
                if (LabelDeleted(s, c)) continue;
                var rawVal = c < series[s].values.Length ? series[s].values[c] : 0;
                var plotVal = stacked ? cumulative[s, c] : rawVal;
                // For percentStacked the displayed value is the original datum,
                // not the normalized 0..100 stack height.
                var textVal = rawVal;
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = stacked ? ClampY(plotVal) : DataToY(plotVal);
                var valuePart = !string.IsNullOrEmpty(dataLabelNumFmt) ? FormatAxisValue(textVal, dataLabelNumFmt)
                    : !string.IsNullOrEmpty(valNumFmt) ? FormatAxisValue(textVal, valNumFmt)
                    : textVal % 1 == 0 ? $"{(int)textVal}" : $"{textVal:0.#}";
                var lparts = new List<string>();
                if (showSerName && s < series.Count) lparts.Add(series[s].name);
                if (showCatName && c < categories.Length) lparts.Add(categories[c]);
                if (showVal) lparts.Add(valuePart);
                var vlabel = string.Join(", ", lparts);
                sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{px:0.#}\" y=\"{py - 6:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"middle\">{HtmlEncode(vlabel)}</text>");
            }

        // Trendlines (parity with bar/line; area previously dropped them — it had
        // no `trendlines` parameter, so a <c:trendline> on an area series was
        // extracted into ChartInfo.Trendlines and then silently discarded. Real
        // PowerPoint overlays the fitted curve on the area fills). Each series is
        // regressed over its 1-based category index; tlMapX converts that index
        // back to the vertex pixel, matching the line renderer.
        if (trendlines != null)
            for (int s = 0; s < series.Count; s++)
            {
                var tl = s < trendlines.Count ? trendlines[s] : null;
                if (tl == null) continue;
                var vals = series[s].values;
                if (vals.Length < 2) continue;
                var lineColor = tl.Color ?? colors[s % colors.Count];
                var xData = new double[vals.Length];
                var yData = new double[vals.Length];
                for (int i = 0; i < vals.Length; i++) { xData[i] = i + 1; yData[i] = vals[i]; }
                Func<double, double> tlMapX = xv => ox + (catCount > 1 ? pw * (xv - 1) / (catCount - 1) : pw / 2.0);
                AppendTrendline(sb, tl, xData, yData, tlMapX, DataToY, lineColor, ox + pw, oy + 12);
            }
    }

    public void RenderRadarChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH, int catLabelFontSize = 0,
        string radarStyle = "filled",
        bool showDataLabels = false, bool showVal = true, bool showSerName = false,
        bool showCatName = false, string? dataLabelNumFmt = null)
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
                // R16-6: OOXML "standard" radar is filled+outline (with markers),
                // same fill as "filled"; only "marker" style stays unfilled.
                var isFilled = radarStyle is "filled" or "standard";
                var fillAttr = isFilled ? $"fill=\"{serColor}\" fill-opacity=\"0.7\"" : "fill=\"none\"";
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
                // Data labels at each vertex (parity with bar/line/area/pie). Placed
                // just outside the vertex along its radial direction so they clear the
                // marker and polygon edge.
                if (showDataLabels)
                for (int c = 0; c < series[s].values.Length && c < catCount; c++)
                {
                    var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
                    var rawVal = series[s].values[c];
                    var rad = rawVal / maxVal * r;
                    var lx = cx + (rad + 10) * Math.Cos(angle);
                    var ly = cy + (rad + 10) * Math.Sin(angle);
                    var valuePart = !string.IsNullOrEmpty(dataLabelNumFmt) ? FormatAxisValue(rawVal, dataLabelNumFmt)
                        : rawVal % 1 == 0 ? $"{(int)rawVal}" : $"{rawVal:0.#}";
                    var lparts = new List<string>();
                    if (showSerName) lparts.Add(series[s].name);
                    if (showCatName && c < categories.Length) lparts.Add(categories[c]);
                    if (showVal) lparts.Add(valuePart);
                    var vlabel = string.Join(", ", lparts);
                    sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"middle\" dominant-baseline=\"middle\">{HtmlEncode(vlabel)}</text>");
                }
            }
        }
        foreach (var frac in new[] { 0.2, 0.4, 0.6, 0.8, 1.0 })
        {
            var val = maxVal * frac;
            var tickLabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{cx + 2:0.#}\" y=\"{cy - r * frac:0.#}\" fill=\"{AxisColor}\" font-size=\"8\"{ValTickWeightAttr} dominant-baseline=\"middle\">{tickLabel}</text>");
        }
        var labelOffset = Math.Max(18, r * 0.15);
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
            var lx = cx + (r + labelOffset) * Math.Cos(angle);
            var ly = cy + (r + labelOffset) * Math.Sin(angle);
            var anchor = Math.Abs(Math.Cos(angle)) < 0.1 ? "middle" : (Math.Cos(angle) > 0 ? "start" : "end");
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"{CatColor}\" font-size=\"{labelSize}\"{CatTickWeightAttr} text-anchor=\"{anchor}\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
        }
    }

    public void RenderBubbleChartSvg(StringBuilder sb, PlotArea plotArea,
        List<(string name, double[] values)> series, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph,
        bool showDataLabels = false, bool showVal = true, bool showSerName = false,
        bool showCatName = false, string? dataLabelNumFmt = null)
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
        var maxSz = allSize.Count > 0 ? allSize.Max() : 1; if (maxSz <= 0) maxSz = 1;
        var bubbleScaleEl = plotArea.Descendants<BubbleScale>().FirstOrDefault();
        var bubbleScale = bubbleScaleEl?.Val?.HasValue == true ? bubbleScaleEl.Val.Value / 100.0 : 1.0;
        var maxRadius = Math.Min(pw, ph) * 0.12 * bubbleScale;
        // c:sizeRepresents — "area" (default): bubble AREA ∝ size (r ∝ √size);
        // "w": bubble WIDTH/diameter ∝ size (r ∝ size, linear). PowerPoint renders the
        // two very differently (in width mode the smallest bubble is a near-invisible
        // dot); the renderer previously always used the area formula.
        var sizeRepEl = plotArea.Descendants<SizeRepresents>().FirstOrDefault();
        var widthMode = sizeRepEl?.Val?.HasValue == true && sizeRepEl.Val.InnerText == "w";

        // Nice axes (round, evenly-spaced ticks) — same approach the scatter
        // renderer uses, so the bubble axes match PowerPoint instead of a raw
        // (max-min)/4 linear division producing fractional ticks. Both X and Y
        // are numeric value axes. The bubble point positions below are mapped
        // through the SAME min/max/step used for the ticks, so bubbles stay
        // aligned with their gridlines.
        double dataMinX = allX.Min(); double dataMaxX = allX.Max();
        bool xIsSmallInteger = allX.All(v => v % 1 == 0) && dataMaxX - dataMinX <= 10;
        double minX; double maxX; double xStep; int xTicks;
        if (xIsSmallInteger && dataMinX > 0)
        {
            // X data is a small positive integer sequence (the common
            // category-index case 1..n). PowerPoint draws integer ticks pinned to
            // the data range with no headroom: 1,2,3,4 for a 1..4 domain.
            minX = dataMinX; maxX = dataMaxX; xStep = 1;
            xTicks = (int)Math.Round(maxX - minX);
            if (xTicks < 1) xTicks = 1;
        }
        else if (dataMinX > 0)
        {
            // Positive non-trivial X: zero-free nice range so ticks stay round
            // rather than spanning 0..max; mirror PowerPoint.
            var (nMaxX, stepX, nX) = ComputeNiceAxisFromMin(Math.Floor(dataMinX), dataMaxX);
            minX = Math.Floor(dataMinX); maxX = nMaxX; xStep = stepX; xTicks = nX;
        }
        else
        {
            var (nMaxX, stepX, nX) = ComputeNiceAxis(dataMaxX);
            minX = 0; maxX = nMaxX; xStep = stepX; xTicks = nX;
        }
        if (maxX <= minX) { maxX = minX + 1; xStep = 1; xTicks = 1; }

        double minY = Math.Min(0, allY.Min());
        var (niceMaxY, tickStepY, nTicksY) = ComputeNiceAxis(allY.Max());
        double maxY = minY >= 0 ? niceMaxY : niceMaxY;
        if (minY >= 0) minY = 0;
        if (maxY <= minY) maxY = minY + 1;

        double MapX(double v) => ox + ((v - minX) / (maxX - minX)) * pw;
        double MapY(double v) => oy + ph - ((v - minY) / (maxY - minY)) * ph;

        if (ShowValGridlines)
        for (int t = 0; t <= nTicksY; t++)
        {
            var gy = MapY(minY + tickStepY * t);
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"{GridlineWidthPx:0.##}\"{ValGridDashAttr}/>");
        }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        for (int s = 0; s < seriesData.Count; s++)
        {
            var (xVals, yVals, sizeVals) = seriesData[s];
            var count = Math.Min(xVals.Length, yVals.Length);
            for (int i = 0; i < count; i++)
            {
                var bx = MapX(xVals[i]);
                var by = MapY(yVals[i]);
                var sz = i < sizeVals.Length ? sizeVals[i] : yVals[i];
                var frac = Math.Max(0, sz) / maxSz;
                var r = widthMode ? frac * maxRadius : Math.Sqrt(frac) * maxRadius + maxRadius * 0.15;
                sb.AppendLine($"        <circle cx=\"{bx:0.#}\" cy=\"{by:0.#}\" r=\"{r:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.6\"/>");
                // Data label — the Y value (same convention as the scatter renderer),
                // placed just outside the bubble to the right so it clears the fill.
                if (showDataLabels)
                {
                    var yv = yVals[i];
                    var valuePart = !string.IsNullOrEmpty(dataLabelNumFmt) ? FormatAxisValue(yv, dataLabelNumFmt)
                        : yv % 1 == 0 ? $"{(int)yv}" : $"{yv:0.#}";
                    var lparts = new List<string>();
                    if (showSerName && s < series.Count) lparts.Add(series[s].name);
                    if (showCatName && i < categories.Length) lparts.Add(categories[i]);
                    if (showVal) lparts.Add(valuePart);
                    var vlabel = string.Join(", ", lparts);
                    sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{bx + r + 4:0.#}\" y=\"{by:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"start\" dominant-baseline=\"middle\">{HtmlEncode(vlabel)}</text>");
                }
            }
        }
        for (int t = 0; t <= xTicks; t++)
        {
            var val = minX + xStep * t;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{MapX(val):0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\"{CatTickWeightAttr} text-anchor=\"middle\">{label}</text>");
        }
        for (int t = 0; t <= nTicksY; t++)
        {
            var val = minY + tickStepY * t;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{MapY(val):0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\"{ValTickWeightAttr} text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    /// <summary>
    /// Scatter (XY) chart. DrawingML scatterChart stores each point as an
    /// xVal/yVal pair (NOT cat/val), so BOTH axes are numeric value axes — the
    /// X axis must reflect the actual numeric X domain, not a 0..n category
    /// index. Reusing the line renderer (category-indexed X with cat labels)
    /// mislabels the X axis (e.g. all "0" or evenly-spaced indices) whenever the
    /// X data isn't a 1..n sequence. This dedicated path maps points through the
    /// real X domain and emits nice numeric X tick labels, matching Excel.
    /// </summary>
    public void RenderScatterChartSvg(StringBuilder sb, PlotArea plotArea,
        List<(string name, double[] values)> series, List<string> colors,
        int ox, int oy, int pw, int ph, List<string>? markerShapes, List<int>? markerSizes,
        List<double>? lineWidths, List<string>? lineDashes, bool markersOnly,
        bool showDataLabels, double? axisMin, double? axisMax, double? majorUnit, string? valNumFmt,
        List<bool>? smooth = null, List<TrendlineInfo?>? trendlines = null, List<ErrorBarInfo?>? errorBars = null,
        List<string?>? markerFillColors = null, List<string?>? markerLineColors = null,
        string? dataLabelNumFmt = null,
        bool showVal = true, bool showSerName = false, bool showCatName = false,
        List<bool>? lineHide = null)
    {
        var scatterSeries = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent?.LocalName == "scatterChart").ToList();

        var allX = new List<double>(); var allY = new List<double>();
        var seriesData = new List<(double[] x, double[] y)>();
        for (int s = 0; s < scatterSeries.Count; s++)
        {
            var ser = scatterSeries[s];
            var xVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "xVal")) ?? [];
            var yVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "yVal")) ?? [];
            seriesData.Add((xVals, yVals));
            allX.AddRange(xVals); allY.AddRange(yVals);
        }
        // Fallback: no scatter ser found (or no cached X) — synthesize an index
        // X domain from the Y series so the chart still renders.
        if (seriesData.Count == 0 || allX.Count == 0)
        {
            seriesData.Clear(); allX.Clear(); allY.Clear();
            foreach (var s in series)
            {
                var xVals = Enumerable.Range(0, s.values.Length).Select(i => (double)i).ToArray();
                seriesData.Add((xVals, s.values));
                allX.AddRange(xVals); allY.AddRange(s.values);
            }
        }
        if (allY.Count == 0) return;

        var dataMinX = allX.Min(); var dataMaxX = allX.Max();
        // Nice X domain: 0-based when all X are non-negative (Excel default),
        // else span the actual min..max. ooxml axis min/max override.
        double minX = axisMin ?? Math.Min(0, dataMinX);
        double maxX = axisMax ?? dataMaxX;
        if (maxX <= minX) maxX = minX + 1;
        double minY = Math.Min(0, allY.Min()); double maxY = allY.Max();
        if (maxY <= minY) maxY = minY + 1;
        var (niceMaxY, tickStepY, nTicksY) = ComputeNiceAxis(maxY);
        if (minY >= 0) { minY = 0; maxY = niceMaxY; }

        // X tick count: honor majorUnit if given, else 4 nice divisions.
        int xTicks = 4;
        double xStep = (maxX - minX) / xTicks;
        if (majorUnit.HasValue && majorUnit.Value > 0)
        {
            xStep = majorUnit.Value;
            xTicks = Math.Max(1, (int)Math.Round((maxX - minX) / xStep));
        }

        double MapX(double v) => ox + ((v - minX) / (maxX - minX)) * pw;
        double MapY(double v) => oy + ph - ((v - minY) / (maxY - minY)) * ph;

        // Gridlines (horizontal, on the Y value axis)
        if (ShowValGridlines)
            for (int t = 0; t <= nTicksY; t++)
            {
                var gy = MapY(minY + tickStepY * t);
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"{GridlineWidthPx:0.##}\"{ValGridDashAttr}/>");
            }
        // Gridlines (vertical, on the X value axis). Scatter has no catAx, so the
        // X-axis majorGridlines are routed into ShowCatGridlines by ExtractChartInfo.
        // Draw at the same X tick positions used for the X labels below.
        if (ShowCatGridlines)
            for (int t = 0; t <= xTicks; t++)
            {
                var gx = MapX(minX + xStep * t);
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
            }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        for (int s = 0; s < seriesData.Count; s++)
        {
            var (xVals, yVals) = seriesData[s];
            var count = Math.Min(xVals.Length, yVals.Length);
            if (count == 0) continue;
            var color = colors[s % colors.Count];
            var pts = new List<(double x, double y, double xv, double yv)>();
            for (int i = 0; i < count; i++)
                pts.Add((MapX(xVals[i]), MapY(yVals[i]), xVals[i], yVals[i]));

            // Connecting line — drawn for lineMarker/smoothMarker scatter styles,
            // suppressed when scatterStyle is marker/none (markersOnly) OR when this
            // series has a:ln/a:noFill (PowerPoint "No line" — markers only). A smooth
            // series uses a Catmull-Rom path (same primitive as the line renderer).
            if (!markersOnly && !(lineHide != null && s < lineHide.Count && lineHide[s]) && pts.Count >= 2)
            {
                var lw = lineWidths != null && s < lineWidths.Count ? lineWidths[s] : 2;
                var dashName = lineDashes != null && s < lineDashes.Count ? lineDashes[s] : "solid";
                var dashAttr = dashName != "solid" ? $" stroke-dasharray=\"{RefLineDashArray(dashName)}\"" : "";
                var isSmooth = smooth != null && s < smooth.Count && smooth[s];
                if (isSmooth)
                {
                    var d = BuildSmoothPath(pts.Select(p => (p.x, p.y)).ToList());
                    sb.AppendLine($"        <path d=\"{d}\" fill=\"none\" stroke=\"{color}\" stroke-width=\"{lw:0.#}\"{dashAttr}/>");
                }
                else
                {
                    var pointStr = string.Join(" ", pts.Select(p => $"{p.x:0.#},{p.y:0.#}"));
                    sb.AppendLine($"        <polyline points=\"{pointStr}\" fill=\"none\" stroke=\"{color}\" stroke-width=\"{lw:0.#}\"{dashAttr}/>");
                }
            }

            // Error bars (vertical, on Y) — shared primitive with the line renderer.
            var ebInfo = errorBars != null && s < errorBars.Count ? errorBars[s] : null;
            if (ebInfo != null)
                AppendErrorBars(sb, pts.Select(p => (p.x, p.y, p.yv)).ToList(), ebInfo, yVals.Take(count).ToArray(), MapY);

            // Trendline — regressed over the REAL X values (xv), so the fitted
            // slope/equation are correct for scatter (unlike the index-based line path).
            var tlInfo = trendlines != null && s < trendlines.Count ? trendlines[s] : null;
            if (tlInfo != null)
            {
                var tlColor = tlInfo.Color ?? color;
                AppendTrendline(sb, tlInfo, pts.Select(p => p.xv).ToArray(), pts.Select(p => p.yv).ToArray(),
                    MapX, MapY, tlColor, ox + pw, oy + 12);
            }

            var shape = markerShapes != null && s < markerShapes.Count ? markerShapes[s] : "circle";
            var mSize = markerSizes != null && s < markerSizes.Count ? markerSizes[s] * 0.6 : 3;
            var mFill = markerFillColors != null && s < markerFillColors.Count ? markerFillColors[s] : null;
            var mStroke = markerLineColors != null && s < markerLineColors.Count ? markerLineColors[s] : null;
            foreach (var p in pts)
            {
                sb.AppendLine($"        {RenderMarkerSvg(shape, p.x, p.y, mSize, mFill ?? color, mStroke ?? color)}");
                if (showDataLabels)
                {
                    // Honor <c:dLbls><c:numFmt> on the data labels (e.g. "$#,##0");
                    // fall back to the value-axis numFmt, then the bare shortcut —
                    // mirrors the line/bubble renderers. Previously hardcoded, so a
                    // currency/percent data label rendered as a bare number.
                    string Fmt(double v) => !string.IsNullOrEmpty(dataLabelNumFmt) ? FormatAxisValue(v, dataLabelNumFmt)
                        : !string.IsNullOrEmpty(valNumFmt) ? FormatAxisValue(v, valNumFmt)
                        : v % 1 == 0 ? $"{(int)v}" : $"{v:0.#}";
                    // Assemble the enabled label parts in PowerPoint's order
                    // (series name, category = X value, value = Y). Previously only
                    // the Y value was emitted, ignoring showSerName/showCatName and
                    // an explicit showVal=0. Mirrors the bubble renderer's lparts.
                    var lparts = new List<string>();
                    if (showSerName && s < series.Count) lparts.Add(series[s].name);
                    if (showCatName) lparts.Add(Fmt(p.xv));
                    if (showVal) lparts.Add(Fmt(p.yv));
                    if (lparts.Count > 0)
                    {
                        var vlabel = string.Join(", ", lparts);
                        sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{p.x:0.#}\" y=\"{p.y - 6:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"middle\">{HtmlEncode(vlabel)}</text>");
                    }
                }
            }
        }

        // Numeric X axis labels (the actual fix: real X domain, not category index)
        for (int t = 0; t <= xTicks; t++)
        {
            var val = minX + xStep * t;
            sb.AppendLine($"        <text x=\"{MapX(val):0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\"{CatTickWeightAttr} text-anchor=\"middle\">{FormatAxisValue(val, valNumFmt)}</text>");
        }
        // Numeric Y axis labels
        for (int t = 0; t <= nTicksY; t++)
        {
            var val = minY + tickStepY * t;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{MapY(val):0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\"{ValTickWeightAttr} text-anchor=\"end\" dominant-baseline=\"middle\">{FormatAxisValue(val, valNumFmt)}</text>");
        }
    }

    public void RenderComboChartSvg(StringBuilder sb, PlotArea plotArea,
        List<(string name, double[] values)> seriesList, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph,
        bool showDataLabels = false, string? dataLabelNumFmt = null,
        double? ooxmlMax = null)
    {
        // Value-label formatter (honors <c:dLbls><c:numFmt>; falls back to a bare value),
        // mirroring the bar/line renderers' data-label formatting.
        string FmtLabel(double v) => !string.IsNullOrEmpty(dataLabelNumFmt) ? FormatAxisValue(v, dataLabelNumFmt)
            : v % 1 == 0 ? $"{(int)v}" : $"{v:0.#}";
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
        // Honor an explicit primary value-axis max (<c:valAx><c:scaling><c:max>): the
        // bar/line renderers pin the top to the entered value (R25); the combo renderer
        // dropped it entirely and always auto-scaled. priNiceMax feeds both the series
        // scaling (axMax) and the tick labels, so overriding it fixes both.
        if (ooxmlMax.HasValue && ooxmlMax.Value > 0) priNiceMax = ooxmlMax.Value;
        var hasSecondary = secondaryValues.Length > 0;
        double secNiceMax = 1;
        if (hasSecondary)
        {
            var secMax = secondaryValues.Max(); if (secMax <= 0) secMax = 1;
            (secNiceMax, _, _) = ComputeNiceAxis(secMax);
        }

        var catCount = Math.Max(categories.Length, seriesList.Max(s => s.values.Length));

        // Primary value-axis horizontal gridlines (<c:valAx><c:majorGridlines/>). The
        // bar/line renderers draw these but the combo path never did, so a combo chart
        // rendered with a blank plot area (PowerPoint shows the major gridlines).
        // Positions match the primary Y-axis labels below; drawn before the series so
        // they sit behind. Gated on ShowValGridlines (synced from info.ValMajorGridlines).
        if (ShowValGridlines)
            for (int t = 0; t <= AxisTickCount; t++)
            {
                var gy = oy + ph - (double)ph * t / AxisTickCount;
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"{GridlineWidthPx:0.##}\"{ValGridDashAttr}/>");
            }

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
                    var bx = ox + c * groupW + gap + bi * barW;
                    sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{oy + ph - barH:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"{FillOpacity(s)}\"/>");
                    if (showDataLabels)
                        sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{bx + barW / 2:0.#}\" y=\"{oy + ph - barH - 3:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"middle\">{HtmlEncode(FmtLabel(val))}</text>");
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
                for (int pi = 0; pi < points.Count; pi++)
                {
                    var parts = points[pi].Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s % colors.Count]}\"/>");
                    if (showDataLabels && pi < seriesList[s].values.Length)
                        sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{parts[0]}\" y=\"{double.Parse(parts[1], System.Globalization.CultureInfo.InvariantCulture) - 6:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"middle\">{HtmlEncode(FmtLabel(seriesList[s].values[pi]))}</text>");
                }
            }
        }
        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (double)pw * c / Math.Max(catCount, 1) + (double)pw / Math.Max(catCount, 1) / 2;
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\"{CatTickWeightAttr} text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        // Primary Y-axis labels (left)
        for (int t = 0; t <= AxisTickCount; t++)
        {
            var val = priNiceMax * t / AxisTickCount;
            var label = FmtValAxis(val);
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{oy + ph - (double)ph * t / AxisTickCount:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\"{ValTickWeightAttr} text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
        // Secondary Y-axis labels (right side, lighter color). R16-5: these used
        // to overlay the primary labels on the left (x=ox+2); placing them at the
        // plot's right edge matches PowerPoint's right-hand secondary axis.
        if (hasSecondary)
        {
            var secFontPx = Math.Max(ValFontPx - 1, CatFontPx);
            for (int t = 0; t <= AxisTickCount; t++)
            {
                var val = secNiceMax * t / AxisTickCount;
                var label = FormatAxisValue(val);
                sb.AppendLine($"        <text x=\"{ox + pw + 4}\" y=\"{oy + ph - (double)ph * t / AxisTickCount:0.#}\" fill=\"{SecondaryAxisColor}\" font-size=\"{secFontPx}\" text-anchor=\"start\" dominant-baseline=\"middle\">{label}</text>");
            }
        }
    }

    /// <summary>Format a VALUE-AXIS tick label, applying the display-units divisor
    /// (<c:dispUnits>) before formatting so a "millions" axis shows 1,2,3 instead
    /// of 1M,2M,3M. Data-label / category-label sites keep calling FormatAxisValue
    /// with the raw value (display units affect the axis only).</summary>
    private string FmtValAxis(double val, string? numFmt = null)
        => FormatAxisValue(ValAxisUnitDivisor != 1.0 ? val / ValAxisUnitDivisor : val, numFmt);

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

        // Extract literal prefix (e.g. "$"). The format code comes from the
        // chart XML (attacker-controlled in a crafted file) and the prefix is
        // emitted into SVG <text>; escape it so a format like "<#,##0" can't
        // inject markup. Value labels/axis ticks are otherwise emitted raw,
        // unlike category labels which are already HtmlEncode'd.
        if (f.Length > 0 && !char.IsDigit(f[0]) && f[0] != '#' && f[0] != '0' && f[0] != '.')
        {
            prefix = HtmlEncode(f[0].ToString());
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
                : val.ToString("N0");  // round, don't truncate ((long) cast truncated 1234.6 -> 1234)
        else
            formatted = decimals > 0
                ? val.ToString($"F{decimals}")
                : (val % 1 == 0 ? $"{(long)val}" : $"{val:0.#}");

        return prefix + formatted + suffix;
    }

    public void RenderStockChartSvg(StringBuilder sb, PlotArea plotArea,
        List<(string name, double[] values)> series, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph, string? upBarColor = null, string? downBarColor = null)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max(); var minVal = allValues.Min();
        if (maxVal <= minVal) maxVal = minVal + 1;
        var range = maxVal - minVal;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        // Use the already-resolved up/down bar colors (ExtractFillColor + themeColors, so
        // schemeClr/sysClr/prstClr all resolve), falling back to the OOXML stock-candlestick
        // defaults (white up / black down) when the bars carry no fill. The previous inline
        // read only handled srgbClr, so a themed candlestick rendered white/black.
        string Hashed(string? c, string fallback)
            => string.IsNullOrEmpty(c) ? fallback : (c.StartsWith('#') ? c : "#" + c);
        var upColor = Hashed(upBarColor, "#FFFFFF");
        var downColor = Hashed(downBarColor, "#000000");

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
        else if (series.Count == 3)
        {
            // R16-4: 3-series stock = hi-lo-close. Render a vertical wick from
            // high to low plus a right-side close tick at each category, instead
            // of falling back to three plain line series.
            var wickColor = downColor == "#000000" ? "#000000" : downColor;
            for (int c = 0; c < catCount; c++)
            {
                var high = c < series[0].values.Length ? series[0].values[c] : 0;
                var low = c < series[1].values.Length ? series[1].values[c] : 0;
                var close = c < series[2].values.Length ? series[2].values[c] : 0;
                var ccx = ox + c * groupW + groupW / 2;
                var yHigh = oy + ph - ((high - minVal) / range) * ph;
                var yLow = oy + ph - ((low - minVal) / range) * ph;
                var yClose = oy + ph - ((close - minVal) / range) * ph;
                var tickW = groupW * 0.25;
                sb.AppendLine($"        <line x1=\"{ccx:0.#}\" y1=\"{yHigh:0.#}\" x2=\"{ccx:0.#}\" y2=\"{yLow:0.#}\" stroke=\"{wickColor}\" stroke-width=\"1.5\"/>");
                sb.AppendLine($"        <line x1=\"{ccx:0.#}\" y1=\"{yClose:0.#}\" x2=\"{ccx + tickW:0.#}\" y2=\"{yClose:0.#}\" stroke=\"{wickColor}\" stroke-width=\"1.5\"/>");
            }
        }
        else { RenderLineChartSvg(sb, series, categories, colors, ox, oy, pw, ph); return; }

        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            sb.AppendLine($"        <text x=\"{ox + c * groupW + groupW / 2:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\"{CatTickWeightAttr} text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        for (int t = 0; t <= 4; t++)
        {
            var val = minVal + range * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{oy + ph - (double)ph * t / 4:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\"{ValTickWeightAttr} text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
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

    // Min-aware nice axis: used only when an explicit non-zero axis min is set
    // and no explicit axis max. PowerPoint derives the tick step from the
    // VISIBLE range [niceMin, dataMax] (not [0, dataMax]) and snaps the top to
    // the smallest niceMin + k*step >= dataMax. e.g. min=50, dataMax=200 →
    // step 20, top 210 (ticks 50,70,...,210). Falls back to the zero-based
    // ComputeNiceAxis when niceMin <= 0 so the common case is byte-identical.
    public static (double niceMax, double tickStep, int nTicks) ComputeNiceAxisFromMin(double niceMin, double dataMax)
    {
        if (niceMin <= 0) return ComputeNiceAxis(dataMax);
        var range = dataMax - niceMin;
        if (range <= 0 || !double.IsFinite(range)) return ComputeNiceAxis(dataMax);
        var mag = Math.Pow(10, Math.Floor(Math.Log10(range)));
        if (!double.IsFinite(mag) || mag == 0) mag = 1;
        var res = range / mag;
        var tickStep = res <= 1.5 ? 0.2 * mag : res <= 4 ? 0.5 * mag : res <= 8 ? 1.0 * mag : 2.0 * mag;
        // smallest tick >= dataMax above the min; add a step of headroom only
        // when dataMax lands exactly on a tick (PowerPoint keeps a gap above the
        // top data point).
        var steps = Math.Ceiling((dataMax - niceMin) / tickStep);
        if (Math.Abs(niceMin + steps * tickStep - dataMax) < 1e-9) steps += 1;
        var niceMax = niceMin + steps * tickStep;
        var nTicks = (int)Math.Round(steps);
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
        // Per-data-point fill overrides for non-pie charts (bar/column). Index =
        // series index; inner dict maps zero-based category idx -> '#'-prefixed
        // hex. Populated from each series' <c:dPt> children. Empty/absent series
        // dicts fall back to the per-series Colors entry (regression-safe).
        public List<Dictionary<int, string>> PerPointColors { get; set; } = [];
        // Per-series set of data-point indices whose label was explicitly deleted
        // (<c:dLbls><c:dLbl><c:idx><c:delete/>); those points show no label even when
        // the series has showVal on. One set per series, aligned to series order.
        public List<HashSet<int>> PerPointDeletedLabels { get; set; } = [];
        public string? Title { get; set; }
        public string TitleFontSize { get; set; } = "10pt";
        public bool TitleBold { get; set; } = true;   // chart titles default to bold
        // Per-run title formatting. PowerPoint renders a chart title with mixed
        // per-run bold/color/size (e.g. one bold-red word + a normal-black word);
        // the single Title/TitleBold/TitleFontColor fields only capture the first
        // run. When the title has runs with differing formatting, TitleRuns holds
        // each run so the render sites can emit styled <span>s. Null/single-run =
        // fall back to the uniform Title string.
        public List<TitleRunInfo>? TitleRuns { get; set; }
        public bool ShowDataLabels { get; set; }
        public bool ShowDataLabelVal { get; set; }
        public bool ShowDataLabelPercent { get; set; }
        public bool ShowDataLabelCatName { get; set; }
        public bool ShowDataLabelSerName { get; set; }
        // <c:dLblPos val="inEnd|outEnd|ctr|inBase"/> — drives where the label sits
        // relative to the bar end. Default "outEnd" matches Office's grouped-bar
        // default; "inEnd" places it inside the bar near its end.
        public string DataLabelPos { get; set; } = "outEnd";
        // True only when <c:dLblPos> is present in the XML (see renderer field).
        public bool HasExplicitDataLabelPos { get; set; }
        public double HoleRatio { get; set; }
        // Pie/doughnut slice explosion as a fraction of radius per data point
        // (index = data point). Empty list = no explosion. Populated from the
        // series-level c:explosion (applies to all points) and/or per-point
        // c:dPt/c:explosion overrides.
        public List<double> Explosions { get; set; } = [];
        // <c:firstSliceAng> — degrees clockwise to rotate the first pie/doughnut
        // slice's start edge from 12 o'clock. 0 = top (default).
        public int FirstSliceAngle { get; set; }
        // Per-series fill opacity parsed from the series spPr <a:alpha>. Index =
        // series index (pie/doughnut: per data point). null = no explicit alpha.
        public List<double?> SeriesFillOpacities { get; set; } = [];
        // Per-series <c:invertIfNegative> (bar/column). Defaults to TRUE when the
        // element is absent (PowerPoint renders negative bars hollow by default);
        // FALSE only when explicitly <c:invertIfNegative val="0"/>. Index = series.
        public List<bool> InvertIfNegative { get; set; } = [];
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
        /// <summary>Value-axis &lt;c:minorUnit&gt;. Drives the minor-gridline count
        /// (majorUnit / minorUnit sub-intervals); null = renderer default.</summary>
        public double? MinorUnit { get; set; }
        public int? GapWidth { get; set; }
        public int? Overlap { get; set; }
        public string? ValAxisTitle { get; set; }
        public int ValAxisTitleFontPx { get; set; } = 9;
        public bool ValAxisTitleBold { get; set; }
        public bool ValAxisTitleItalic { get; set; }
        // Explicit axis-title run color ('#'-prefixed CSS, or null = use AxisColor).
        // PowerPoint honors a solidFill on the axis-title run; previously dropped.
        public string? ValAxisTitleColor { get; set; }
        public string? CatAxisTitle { get; set; }
        public int CatAxisTitleFontPx { get; set; } = 9;
        public bool CatAxisTitleBold { get; set; }
        public bool CatAxisTitleItalic { get; set; }
        public string? CatAxisTitleColor { get; set; }
        // Secondary value axis title (combo right-side axis; also the bubble/scatter
        // Y axis, which is the 2nd valAx rather than a catAx).
        public string? SecondaryValAxisTitle { get; set; }
        public int SecondaryValAxisTitleFontPx { get; set; } = 9;
        public bool SecondaryValAxisTitleBold { get; set; }
        public bool SecondaryValAxisTitleItalic { get; set; }
        public string? SecondaryValAxisTitleColor { get; set; }
        public string? PlotFillColor { get; set; }
        public string? ChartFillColor { get; set; }
        /// <summary>Plot-area &lt;c:spPr&gt;&lt;a:ln&gt; outline color (hex, no #). Null = no border.</summary>
        public string? PlotBorderColor { get; set; }
        /// <summary>Plot-area outline width in EMU (a:ln/@w). Null defaults to PowerPoint's ~0.75pt.</summary>
        public long? PlotBorderWidthEmu { get; set; }
        /// <summary>Chart-area &lt;c:spPr&gt;&lt;a:ln&gt; outline color (hex, no #). Null = no border.</summary>
        public string? ChartBorderColor { get; set; }
        /// <summary>Chart-area outline width in EMU. Null defaults to PowerPoint's ~0.75pt.</summary>
        public long? ChartBorderWidthEmu { get; set; }
        public bool HasLegend { get; set; }
        /// <summary>Series indices whose &lt;c:legendEntry&gt;&lt;c:delete val="1"/&gt;
        /// hides the legend swatch+label while the series still plots. Empty
        /// for charts with no deleted entries (the common case).</summary>
        public HashSet<int> DeletedLegendEntries { get; set; } = new();
        /// <summary>#7f: OOXML c:legendPos InnerText — "r" (right, ECMA-376
        /// CT_LegendPos default), "b" (bottom), "t" (top), "l" (left),
        /// "tr" (top-right). Default is "r" to match the schema default that
        /// real PowerPoint applies when &lt;c:legendPos&gt; is absent. Rendering
        /// adapts the wrapper layout to each position.</summary>
        public string LegendPos { get; set; } = "r";
        public string LegendFontSize { get; set; } = "8pt";
        public string? LegendFontColor { get; set; }
        public bool LegendFontBold { get; set; }
        public int ValFontPx { get; set; } = 9;
        public string? ValFontColor { get; set; }
        public bool ValFontBold { get; set; }
        public bool ValFontItalic { get; set; }
        public int CatFontPx { get; set; } = 9;
        public string? CatFontColor { get; set; }
        public bool CatFontBold { get; set; }
        public bool CatFontItalic { get; set; }
        /// <summary>Category-axis tick-label rotation in degrees, read from
        /// &lt;c:catAx&gt;&lt;c:txPr&gt;&lt;a:bodyPr rot="..."/&gt; (OOXML rot is
        /// 1/60000 degree). Null = no rotation (labels horizontal, default).</summary>
        public int? CatAxisLabelRotationDeg { get; set; }
        /// <summary>Value-axis tick-label rotation in degrees (analogous to
        /// CatAxisLabelRotationDeg). Null = no rotation.</summary>
        public int? ValAxisLabelRotationDeg { get; set; }
        public string? ValNumFmt { get; set; }
        /// <summary>Value-axis display units (&lt;c:dispUnits&gt;&lt;c:builtInUnit&gt;):
        /// the divisor PowerPoint applies to every value-axis tick label (e.g.
        /// millions → 1e6, so 1,000,000 shows as "1"). 1.0 = no scaling.</summary>
        public double ValueAxisUnitDivisor { get; set; } = 1.0;
        /// <summary>The rotated annotation drawn beside the value axis when display
        /// units are set and &lt;c:dispUnitsLbl&gt; is present (e.g. "Millions").
        /// Null = no label.</summary>
        public string? ValueAxisUnitLabel { get; set; }
        /// <summary>Format code from &lt;c:dLbls&gt;&lt;c:numFmt&gt; — applied to data
        /// labels (overrides the value-axis ValNumFmt for label text).</summary>
        public string? DataLabelsNumFmt { get; set; }
        public string? TitleFontColor { get; set; }
        public string? GridlineColor { get; set; }
        /// <summary>Value-axis major-gridline OOXML dash name (&lt;a:prstDash val="..."/&gt;,
        /// e.g. "dash"). Null when absent or "solid".</summary>
        public string? GridlineDash { get; set; }
        /// <summary>Value-axis major-gridline line width (&lt;a:ln w="..."/&gt; EMU).
        /// Null = use the renderer's default thin gridline.</summary>
        public long? GridlineWidthEmu { get; set; }
        /// <summary>True when the value axis has &lt;c:majorGridlines&gt; (horizontal gridlines).</summary>
        public bool ValMajorGridlines { get; set; }
        /// <summary>True when the category axis has &lt;c:majorGridlines&gt; (vertical gridlines).</summary>
        public bool CatMajorGridlines { get; set; }
        /// <summary>True when the value axis has &lt;c:minorGridlines&gt; (fainter sub-interval gridlines).</summary>
        public bool ValMinorGridlines { get; set; }
        /// <summary>True when the category axis has &lt;c:minorGridlines&gt;.</summary>
        public bool CatMinorGridlines { get; set; }
        /// <summary>False when the value axis is deleted (&lt;c:delete val="1"/&gt;). Default true.</summary>
        public bool ValAxisVisible { get; set; } = true;
        public bool ValTickLabelsHidden { get; set; }
        public bool CatTickLabelsHidden { get; set; }
        /// <summary>False when the category axis is deleted (&lt;c:delete val="1"/&gt;). Default true.</summary>
        public bool CatAxisVisible { get; set; } = true;
        public string? AxisLineColor { get; set; }
        /// <summary>Value axis &lt;c:majorTickMark val="..."/&gt; ("out"/"in"/"cross"/"none"). Null when absent.</summary>
        public string? ValMajorTickMark { get; set; }
        /// <summary>Category axis &lt;c:majorTickMark val="..."/&gt; ("out"/"in"/"cross"/"none"). Null when absent.</summary>
        public string? CatMajorTickMark { get; set; }
        public int CatTickLabelSkip { get; set; } = 1;
        public int DataLabelFontPx { get; set; } = 8;
        /// <summary>Explicit data-label text color from &lt;c:dLbls&gt;&lt;c:txPr&gt;
        /// (bare-hex resolved via theme), or null to use the theme text color.</summary>
        public string? DataLabelFontColor { get; set; }
        public bool DataLabelBold { get; set; }
        public bool DataLabelItalic { get; set; }
        /// <summary>Reference-line overlays (horizontal dashed lines at constant values).
        /// Filled by ExtractChartInfo from any ref-line-only LineChart in the plot area.</summary>
        public List<(string Name, double Value, string Color, double WidthPt, string Dash)> ReferenceLines { get; set; } = [];

        // --- Marker shapes per series (circle, diamond, square, triangle, star, x, plus, dash, dot, none) ---
        public List<string> MarkerShapes { get; set; } = [];
        public List<int> MarkerSizes { get; set; } = [];
        // Per-series marker fill / border (from <c:marker><c:spPr>). null => use
        // the series color (PowerPoint's default: solid series-colored marker).
        public List<string?> MarkerFillColors { get; set; } = [];
        public List<string?> MarkerLineColors { get; set; } = [];

        // --- Smooth line (cubic spline) per series ---
        public List<bool> Smooth { get; set; } = [];

        // --- Dash pattern per series (solid, dash, dot, dashDot, lgDash, etc.) ---
        public List<string> LineDashes { get; set; } = [];

        // --- Line width per series (in points, from a:ln w="...") ---
        public List<double> LineWidths { get; set; } = [];

        // --- Per-series "no connecting line" flag (a:ln/a:noFill on the series spPr) ---
        // PowerPoint's "Format Data Series -> Line -> No line": the series shows markers
        // only, no polyline. Distinct from the chart-wide scatterStyle=marker.
        public List<bool> SeriesLineHide { get; set; } = [];

        // --- Axis features ---
        public double? LogBase { get; set; }
        public bool IsReversed { get; set; }       // value axis maxMin
        public bool IsCatReversed { get; set; }    // category axis maxMin

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

        // R16c: scatterStyle="marker" (not lineMarker/smoothMarker) = dots only,
        // no connecting line. Suppresses the polyline in the line/scatter renderer.
        public bool ScatterMarkersOnly { get; set; }

        // --- Radar style (standard, marker, filled) ---
        public string RadarStyle { get; set; } = "filled";

        // --- Trendlines per series ---
        public List<TrendlineInfo?> Trendlines { get; set; } = [];

        // --- Error bars per series ---
        public List<ErrorBarInfo?> ErrorBars { get; set; } = [];
    }

    /// <summary>One run of a chart title's rich text, for per-run styled rendering.</summary>
    public class TitleRunInfo
    {
        public string Text { get; set; } = "";
        public bool? Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        /// <summary>'#'-prefixed CSS color, or null to inherit the title default.</summary>
        public string? Color { get; set; }
        public double? FontSizePt { get; set; }
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
        public bool NoEndCap { get; set; }
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
    public static ChartInfo ExtractChartInfo(OpenXmlElement plotArea, OpenXmlElement? chart,
        Dictionary<string, string>? themeColors = null)
    {
        var info = new ChartInfo();
        info.PlotArea = plotArea as PlotArea;
        if (info.PlotArea == null) return info;

        // Chart type, categories, series
        info.ChartType = ChartHelper.DetectChartType(info.PlotArea) ?? "column";
        info.Categories = ChartHelper.ReadCategories(info.PlotArea) ?? [];
        info.Series = ChartHelper.ReadAllSeries(info.PlotArea);
        info.ReferenceLines = ChartHelper.ReadReferenceLines(info.PlotArea, themeColors);

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
        // Gather ser elements across ALL chart-type groups (in the same document order
        // ReadAllSeries uses to build info.Series), not just the first group. A combo
        // chart has a barChart AND a lineChart group; taking only chartTypeEl's (first
        // group's) ser dropped the line-group series, so they fell through to fallback
        // colors. Each series' color is then resolved per-series by its parent group's
        // type (line/scatter → stroke color, others → fill) inside ExtractColors.
        var serElements = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent != null
                && (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart")))
            .Cast<OpenXmlElement>()
            .ToList();
        // Pie/doughnut varyColors: default true (vary by point) when the element is
        // absent; explicit val="0"/"false" makes the pie monochrome (series color).
        var pieVcEl = isPieType ? chartTypeEl?.Elements().FirstOrDefault(e => e.LocalName == "varyColors") : null;
        var pieVcVal = pieVcEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        var pieVaryColors = !(pieVcVal is "0" or "false");
        info.Colors = ExtractColors(serElements, info.Series, isPieType, info.ChartType, themeColors, pieVaryColors);
        // Per-data-point fill overrides (c:dPt) for non-pie charts. Pie/doughnut
        // already fold dPt into per-point Colors above, so only collect these for
        // the non-pie case where Colors is per-series.
        if (!isPieType)
            info.PerPointColors = ExtractPerPointColors(serElements, themeColors);
        // Per-point deleted-label overrides (read for all chart types, pie included).
        info.PerPointDeletedLabels = ExtractDeletedLabels(serElements);

        // <c:varyColors val="1"/> on a single-series NON-pie chart colors each data
        // point from the theme accent palette (PowerPoint "vary colors by point") —
        // but ONLY when the series has no explicit fill (an explicit series fill wins,
        // keeping the bars monochrome). Seed PerPointColors, which BarFill consults;
        // explicit dPt entries already present are preserved (not overwritten).
        if (!isPieType && info.Series.Count == 1 && serElements.Count == 1)
        {
            var vcEl = chartTypeEl?.Elements().FirstOrDefault(e => e.LocalName == "varyColors");
            var vcVal = vcEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            var varyOn = vcEl != null && vcVal is null or "1" or "true";
            var serSpPr = serElements[0].Elements().FirstOrDefault(e => e.LocalName == "spPr");
            var hasExplicitFill = ExtractFillColor(serSpPr, themeColors) != null;
            if (varyOn && !hasExplicitFill)
            {
                // themeColors is null on the xlsx/docx preview paths (only pptx
                // extracts a theme) — fall back to the static palette there.
                var palette = themeColors == null
                    ? FallbackColors
                    : new[] { "accent1", "accent2", "accent3", "accent4", "accent5", "accent6" }
                        .Where(k => themeColors.ContainsKey(k)).Select(k => $"#{themeColors[k]}").ToArray();
                if (palette.Length == 0) palette = FallbackColors;
                while (info.PerPointColors.Count < 1) info.PerPointColors.Add([]);
                var catN = info.Categories.Count();
                for (int c = 0; c < catN; c++)
                    if (!info.PerPointColors[0].ContainsKey(c))
                        info.PerPointColors[0][c] = palette[c % palette.Length];
            }
        }

        // Title
        var titleEl = chart?.Elements().FirstOrDefault(e => e.LocalName == "title");
        if (titleEl != null)
        {
            var runEls = titleEl.Descendants<Drawing.Run>().ToList();
            info.Title = string.Join("", runEls.Select(r => r.GetFirstChild<Drawing.Text>()?.Text).Where(t => t != null));
            // Capture per-run formatting so a mixed-format title (e.g. a bold word
            // + a normal word) renders with per-run <span>s instead of collapsing
            // to the first run's style. Only kept when >1 run carries text.
            var perRun = new List<TitleRunInfo>();
            foreach (var r in runEls)
            {
                var txt = r.GetFirstChild<Drawing.Text>()?.Text;
                if (txt == null) continue;
                var rp = r.GetFirstChild<Drawing.RunProperties>();
                var c = ExtractFontColor(rp, themeColors);
                perRun.Add(new TitleRunInfo
                {
                    Text = txt,
                    Bold = rp?.Bold?.HasValue == true ? rp.Bold.Value : null,
                    Italic = rp?.Italic?.Value == true,
                    Underline = rp?.Underline?.HasValue == true && rp.Underline.Value != Drawing.TextUnderlineValues.None,
                    Color = c != null ? CssHexColor(c) : null,
                    FontSizePt = rp?.FontSize?.HasValue == true ? rp.FontSize.Value / 100.0 : null,
                });
            }
            if (perRun.Count > 1) info.TitleRuns = perRun;
            var titleRPr = titleEl.Descendants<Drawing.RunProperties>().FirstOrDefault();
            if (titleRPr?.FontSize?.HasValue == true)
                info.TitleFontSize = $"{titleRPr.FontSize.Value / 100.0:0.##}pt";
            info.TitleFontColor = ExtractFontColor(titleRPr, themeColors);
            // Chart title bold: default true, but honor an explicit b="0" (run rPr or the
            // paragraph defRPr). The renderer previously hardcoded font-weight:bold, so a
            // title set to non-bold still rendered bold. Mirrors the axis-title bold path.
            var titleDefRPr = titleEl.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
            if (titleRPr?.Bold?.HasValue == true) info.TitleBold = titleRPr.Bold.Value;
            else if (titleDefRPr?.Bold?.HasValue == true) info.TitleBold = titleDefRPr.Bold.Value;
        }

        // Data labels
        var dLbls = chartTypeEl?.Elements().FirstOrDefault(e => e.LocalName == "dLbls")
            ?? plotArea.Descendants().FirstOrDefault(e => e.LocalName == "dLbls");
        if (dLbls != null)
        {
            // CT_Boolean's val attribute defaults to true, so a bare <c:showVal/> (no val)
            // means ON — PowerPoint emits this form when labels are enabled via the UI.
            // The old `== "1"` check treated bare/`true` as OFF, suppressing all labels.
            bool IsOn(string name) => dLbls.Elements().Any(e =>
            {
                if (e.LocalName != name) return false;
                var v = e.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                return v is null or "" or "1" or "true";
            });
            info.ShowDataLabelVal = IsOn("showVal");
            info.ShowDataLabelPercent = IsOn("showPercent");
            info.ShowDataLabelCatName = IsOn("showCatName");
            info.ShowDataLabelSerName = IsOn("showSerName");
            info.ShowDataLabels = info.ShowDataLabelVal || info.ShowDataLabelPercent || info.ShowDataLabelCatName || info.ShowDataLabelSerName;
            // <c:dLblPos> — inEnd (inside, near end) vs outEnd (beyond end) etc.
            // Office places insideEnd labels within the bar and outsideEnd just
            // past the bar tip; ignoring it made both positions identical.
            var dLblPosEl = dLbls.Elements().FirstOrDefault(e => e.LocalName == "dLblPos");
            var dLblPosVal = dLblPosEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (!string.IsNullOrEmpty(dLblPosVal)) { info.DataLabelPos = dLblPosVal!; info.HasExplicitDataLabelPos = true; }
            // <c:numFmt formatCode="#,##0"> inside dLbls formats the label text
            // (e.g. grouping separators). Independent of the value-axis numFmt.
            var dLblNumFmtEl = dLbls.Elements().FirstOrDefault(e => e.LocalName == "numFmt");
            var dLblFmtCode = dLblNumFmtEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "formatCode").Value;
            if (!string.IsNullOrEmpty(dLblFmtCode)) info.DataLabelsNumFmt = dLblFmtCode;
        }

        // Doughnut hole size
        if (info.ChartType.Contains("doughnut"))
        {
            var holeSizeEl = chartTypeEl?.Elements().FirstOrDefault(e => e.LocalName == "holeSize");
            var holeSizeVal = holeSizeEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            info.HoleRatio = (holeSizeVal != null && int.TryParse(holeSizeVal, out var hs) ? hs : 10) / 100.0; // OOXML spec default: 10%
        }

        // Pie/doughnut slice explosion. Series-level c:explosion applies to every
        // data point; per-point c:dPt/c:explosion overrides a single slice. Values
        // are percentages of the pie radius (PowerPoint: 20 → 20% of radius).
        if (isPieType && serElements.Count > 0)
        {
            var pieSer = serElements[0];
            double serExpl = 0;
            var serExplEl = pieSer.Elements().FirstOrDefault(e => e.LocalName == "explosion");
            if (serExplEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value is string sev
                && double.TryParse(sev, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var sevd))
                serExpl = sevd / 100.0;
            var ptCount = info.Series.Count > 0 ? info.Series[0].values.Length : info.Categories.Length;
            if (ptCount > 0 && (serExpl > 0 || pieSer.Elements().Any(e => e.LocalName == "dPt")))
            {
                var expl = new List<double>();
                for (int i = 0; i < ptCount; i++) expl.Add(serExpl);
                foreach (var dPt in pieSer.Elements().Where(e => e.LocalName == "dPt"))
                {
                    var idxEl = dPt.Elements().FirstOrDefault(e => e.LocalName == "idx");
                    var dExplEl = dPt.Elements().FirstOrDefault(e => e.LocalName == "explosion");
                    if (idxEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value is string ivs
                        && int.TryParse(ivs, out var idx) && idx >= 0 && idx < expl.Count
                        && dExplEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value is string dev
                        && double.TryParse(dev, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var devd))
                        expl[idx] = devd / 100.0;
                }
                if (expl.Any(e => e > 0)) info.Explosions = expl;
            }
        }

        // <c:firstSliceAng> — rotate the pie/doughnut start angle clockwise.
        if (isPieType)
        {
            var fsaEl = chartTypeEl?.Elements().FirstOrDefault(e => e.LocalName == "firstSliceAng");
            if (fsaEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value is string fsav
                && int.TryParse(fsav, out var fsa))
                info.FirstSliceAngle = ((fsa % 360) + 360) % 360;
        }

        // Per-series fill opacity from <a:solidFill><a:alpha val="…"/>. Pie/doughnut
        // alpha lives on per-point dPt spPr; other charts on the series spPr.
        info.SeriesFillOpacities = ExtractFillOpacities(serElements, info.Series, isPieType);

        // Per-series <c:invertIfNegative>. Default TRUE when the element is
        // absent (PowerPoint renders negative bars hollow by default); FALSE
        // only when explicitly val="0". Index aligns with info.Series order.
        info.InvertIfNegative = ExtractInvertIfNegative(serElements, info.Series.Count);

        // Axis info
        var valAxes = plotArea.Elements().Where(e => e.LocalName == "valAx").ToList();
        var valAxis = valAxes.FirstOrDefault();
        var catAxis = plotArea.Elements().FirstOrDefault(e => e.LocalName == "catAx");

        // A second value axis carries the secondary (combo right-side) title, or —
        // for bubble/scatter charts that have no catAx — the Y-axis title. Emit it
        // so it isn't silently dropped.
        var secondaryValAxis = valAxes.Count >= 2 ? valAxes[1] : null;
        if (secondaryValAxis != null)
        {
            var secTitleEl = secondaryValAxis.Elements().FirstOrDefault(e => e.LocalName == "title");
            info.SecondaryValAxisTitle = secTitleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
            var secTitleRPr = secTitleEl?.Descendants<Drawing.RunProperties>().FirstOrDefault();
            if (secTitleRPr?.FontSize?.HasValue == true)
                info.SecondaryValAxisTitleFontPx = (int)(secTitleRPr.FontSize.Value / 100.0);
            if (secTitleRPr?.Bold?.Value == true)
                info.SecondaryValAxisTitleBold = true;
            if (secTitleRPr?.Italic?.Value == true)
                info.SecondaryValAxisTitleItalic = true;
            var secTitleColor = ExtractFontColor(secTitleRPr, themeColors);
            if (secTitleColor != null) info.SecondaryValAxisTitleColor = CssHexColor(secTitleColor);
        }

        if (valAxis != null)
        {
            var valTitleEl = valAxis.Elements().FirstOrDefault(e => e.LocalName == "title");
            info.ValAxisTitle = valTitleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
            var valTitleRPr = valTitleEl?.Descendants<Drawing.RunProperties>().FirstOrDefault();
            if (valTitleRPr?.FontSize?.HasValue == true)
                info.ValAxisTitleFontPx = (int)(valTitleRPr.FontSize.Value / 100.0);
            if (valTitleRPr?.Bold?.Value == true)
                info.ValAxisTitleBold = true;
            if (valTitleRPr?.Italic?.Value == true)
                info.ValAxisTitleItalic = true;
            var valTitleColor = ExtractFontColor(valTitleRPr, themeColors);
            if (valTitleColor != null) info.ValAxisTitleColor = CssHexColor(valTitleColor);
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
            var minorUnit = valAxis.Elements().FirstOrDefault(e => e.LocalName == "minorUnit");
            if (minorUnit != null && double.TryParse(minorUnit.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var mnu) && mnu > 0)
                info.MinorUnit = mnu;

            // Display units (<c:dispUnits><c:builtInUnit val="millions"/>): PowerPoint
            // divides every value-axis tick by the unit and (when <c:dispUnitsLbl> is
            // present) draws a rotated unit-name annotation beside the axis.
            var dispUnits = valAxis.Elements().FirstOrDefault(e => e.LocalName == "dispUnits");
            if (dispUnits != null)
            {
                var builtIn = dispUnits.Elements().FirstOrDefault(e => e.LocalName == "builtInUnit")?
                    .GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                var (div, name) = builtIn switch
                {
                    "hundreds" => (1e2, "Hundreds"),
                    "thousands" => (1e3, "Thousands"),
                    "tenThousands" => (1e4, "Ten Thousands"),
                    "hundredThousands" => (1e5, "Hundred Thousands"),
                    "millions" => (1e6, "Millions"),
                    "tenMillions" => (1e7, "Ten Millions"),
                    "hundredMillions" => (1e8, "Hundred Millions"),
                    "billions" => (1e9, "Billions"),
                    "trillions" => (1e12, "Trillions"),
                    _ => (1.0, null as string),
                };
                if (div != 1.0)
                {
                    info.ValueAxisUnitDivisor = div;
                    // The annotation is shown only when <c:dispUnitsLbl> is present
                    // (PowerPoint's "Show display units label on chart" toggle).
                    if (dispUnits.Elements().Any(e => e.LocalName == "dispUnitsLbl"))
                        info.ValueAxisUnitLabel = name;
                }
            }

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
            info.ValFontColor = ExtractFontColor(valDefRPr, themeColors);
            if (valDefRPr?.Bold?.HasValue == true) info.ValFontBold = valDefRPr.Bold.Value;
            if (valDefRPr?.Italic?.HasValue == true) info.ValFontItalic = valDefRPr.Italic.Value;
            info.ValAxisLabelRotationDeg = ExtractAxisLabelRotationDeg(valTxPr);

            // Gridline color
            var majorGridlines = valAxis.Elements().FirstOrDefault(e => e.LocalName == "majorGridlines");
            info.ValMajorGridlines = majorGridlines != null;
            info.ValMinorGridlines = valAxis.Elements().Any(e => e.LocalName == "minorGridlines");
            var gridSpPr = majorGridlines?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
            info.GridlineColor = ExtractLineColor(gridSpPr, themeColors);
            // Value-axis major-gridline dash style (<a:ln><a:prstDash val="dash"/>).
            var gridLnEl = gridSpPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
            var gridDashEl = gridLnEl?.Elements().FirstOrDefault(e => e.LocalName == "prstDash");
            var gridDashVal = gridDashEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (!string.IsNullOrEmpty(gridDashVal) && gridDashVal != "solid")
                info.GridlineDash = gridDashVal;
            // Value-axis major-gridline width (<a:ln w="EMU"/>). Without this the
            // gridline rendered at a fixed 0.5px regardless of an explicit thick width.
            var gridWidthStr = gridLnEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value;
            if (long.TryParse(gridWidthStr, out var gwEmu) && gwEmu > 0)
                info.GridlineWidthEmu = gwEmu;

            // BUG4(R25): <c:delete val="1"/> hides the axis (ticks + gridlines).
            var valDeleteEl = valAxis.Elements().FirstOrDefault(e => e.LocalName == "delete");
            var valDelVal = valDeleteEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            info.ValAxisVisible = valDelVal != "1";
            // <c:tickLblPos val="none"/>: hide value-axis labels (keep line/gridlines).
            var valTlpEl = valAxis.Elements().FirstOrDefault(e => e.LocalName == "tickLblPos");
            info.ValTickLabelsHidden = valTlpEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == "none";

            // Axis line color
            var valSpPr = valAxis.Elements().FirstOrDefault(e => e.LocalName == "spPr");
            info.AxisLineColor = ExtractLineColor(valSpPr, themeColors);

            // Major tick marks (short perpendicular lines at each major label)
            var valTickEl = valAxis.Elements().FirstOrDefault(e => e.LocalName == "majorTickMark");
            info.ValMajorTickMark = valTickEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;

            // Value axis number format (e.g. "$#,##0")
            var numFmtEl = valAxis.Elements().FirstOrDefault(e => e.LocalName == "numFmt");
            var fmtCode = numFmtEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "formatCode").Value;
            if (!string.IsNullOrEmpty(fmtCode) && fmtCode != "General")
                info.ValNumFmt = fmtCode;
        }
        // Scatter/bubble charts have NO catAx — they use two valAx (axPos="b" = X
        // axis, axPos="l" = Y axis). The bottom valAx carries the X-axis
        // majorGridlines (vertical gridlines). PowerPoint draws them; mirror that
        // by routing the bottom valAx's majorGridlines into CatMajorGridlines.
        if (catAxis == null && valAxes.Count >= 2)
        {
            var bottomValAx = valAxes.FirstOrDefault(va =>
                va.Elements().FirstOrDefault(e => e.LocalName == "axPos")
                    ?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == "b");
            if (bottomValAx != null)
            {
                info.CatMajorGridlines = bottomValAx.Elements().Any(e => e.LocalName == "majorGridlines");
                info.CatMinorGridlines = bottomValAx.Elements().Any(e => e.LocalName == "minorGridlines");
                var bTickEl = bottomValAx.Elements().FirstOrDefault(e => e.LocalName == "majorTickMark");
                info.CatMajorTickMark = bTickEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                var bDeleteEl = bottomValAx.Elements().FirstOrDefault(e => e.LocalName == "delete");
                info.CatAxisVisible = bDeleteEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value != "1";
            }
        }
        if (catAxis != null)
        {
            // Category axis orientation (<c:catAx><c:scaling><c:orientation val="maxMin"/>):
            // reverses the category order. The value-axis equivalent (IsReversed) was read
            // above; the catAx one was dropped, so a reversed category axis rendered forward.
            var catScaling = catAxis.Elements().FirstOrDefault(e => e.LocalName == "scaling");
            var catOrientEl = catScaling?.Elements().FirstOrDefault(e => e.LocalName == "orientation");
            var catOrientVal = catOrientEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            info.IsCatReversed = catOrientVal == "maxMin";
            info.CatMajorGridlines = catAxis.Elements().Any(e => e.LocalName == "majorGridlines");
            info.CatMinorGridlines = catAxis.Elements().Any(e => e.LocalName == "minorGridlines");
            var catTickEl = catAxis.Elements().FirstOrDefault(e => e.LocalName == "majorTickMark");
            info.CatMajorTickMark = catTickEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            // tickLblSkip: thin category labels to every Nth (read but never rendered before).
            var catLblSkipEl = catAxis.Elements().FirstOrDefault(e => e.LocalName == "tickLblSkip");
            if (catLblSkipEl != null
                && int.TryParse(catLblSkipEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var catLblSkip)
                && catLblSkip > 1)
                info.CatTickLabelSkip = catLblSkip;
            var catDeleteEl = catAxis.Elements().FirstOrDefault(e => e.LocalName == "delete");
            var catDelVal = catDeleteEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            info.CatAxisVisible = catDelVal != "1";
            // <c:tickLblPos val="none"/>: hide category-axis labels (keep line/gridlines).
            var catTlpEl = catAxis.Elements().FirstOrDefault(e => e.LocalName == "tickLblPos");
            info.CatTickLabelsHidden = catTlpEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == "none";
            var catTitleEl = catAxis.Elements().FirstOrDefault(e => e.LocalName == "title");
            info.CatAxisTitle = catTitleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
            var catTitleRPr = catTitleEl?.Descendants<Drawing.RunProperties>().FirstOrDefault();
            if (catTitleRPr?.FontSize?.HasValue == true)
                info.CatAxisTitleFontPx = (int)(catTitleRPr.FontSize.Value / 100.0);
            if (catTitleRPr?.Bold?.Value == true)
                info.CatAxisTitleBold = true;
            if (catTitleRPr?.Italic?.Value == true)
                info.CatAxisTitleItalic = true;
            var catTitleColor = ExtractFontColor(catTitleRPr, themeColors);
            if (catTitleColor != null) info.CatAxisTitleColor = CssHexColor(catTitleColor);
            // Use txPr > defRPr for tick label font (not title's RunProperties)
            var catTxPr = catAxis.Elements().FirstOrDefault(e => e.LocalName == "txPr");
            var catDefRPr = catTxPr?.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
            if (catDefRPr?.FontSize?.HasValue == true)
                info.CatFontPx = (int)(catDefRPr.FontSize.Value / 100.0);
            info.CatFontColor = ExtractFontColor(catDefRPr, themeColors);
            if (catDefRPr?.Bold?.HasValue == true) info.CatFontBold = catDefRPr.Bold.Value;
            if (catDefRPr?.Italic?.HasValue == true) info.CatFontItalic = catDefRPr.Italic.Value;
            info.CatAxisLabelRotationDeg = ExtractAxisLabelRotationDeg(catTxPr);
        }

        // Data label font size
        if (dLbls != null)
        {
            var dLblDefRPr = dLbls.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
            var dLblFontSize = dLblDefRPr?.FontSize ?? dLbls.Descendants<Drawing.RunProperties>().FirstOrDefault()?.FontSize;
            if (dLblFontSize?.HasValue == true)
                info.DataLabelFontPx = (int)(dLblFontSize.Value / 100.0);
            // Explicit data-label color (<c:txPr>…<a:solidFill>); resolves schemeClr
            // through the theme. PowerPoint honors it; previously dropped → theme text.
            var dLblColor = ExtractFontColor(dLblDefRPr, themeColors);
            if (dLblColor != null) info.DataLabelFontColor = CssHexColor(dLblColor);
            if (dLblDefRPr?.Bold?.HasValue == true) info.DataLabelBold = dLblDefRPr.Bold.Value;
            if (dLblDefRPr?.Italic?.HasValue == true) info.DataLabelItalic = dLblDefRPr.Italic.Value;
        }

        // Gap width
        var gapWidthEl = plotArea.Descendants().FirstOrDefault(e => e.LocalName == "gapWidth");
        if (gapWidthEl != null)
        {
            var gv = gapWidthEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (gv != null && int.TryParse(gv, out var gw)) info.GapWidth = gw;
        }

        // Overlap (clustered bar/column: percentage two adjacent series bars
        // overlap; 100 = fully overlapping, 0 = touching, negative = gap).
        var overlapEl = plotArea.Descendants().FirstOrDefault(e => e.LocalName == "overlap");
        if (overlapEl != null)
        {
            var ov = overlapEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (ov != null && int.TryParse(ov, out var ow)) info.Overlap = ow;
        }

        // Plot / chart fill
        var plotSpPr = plotArea.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        info.PlotFillColor = ExtractFillColor(plotSpPr, themeColors);
        info.PlotBorderColor = ExtractLineColor(plotSpPr, themeColors);
        info.PlotBorderWidthEmu = ExtractLineWidthEmu(plotSpPr);
        var chartSpPr = chart?.Parent?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        info.ChartFillColor = ExtractFillColor(chartSpPr, themeColors);
        info.ChartBorderColor = ExtractLineColor(chartSpPr, themeColors);
        info.ChartBorderWidthEmu = ExtractLineWidthEmu(chartSpPr);

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
            info.LegendFontColor = ExtractFontColor(legendRPr, themeColors);
            // Legend font bold (<c:legend><c:txPr>…<a:rPr b="1"/> or defRPr): the renderer
            // emitted size+color but never font-weight, so a bold legend rendered normal.
            // Mirrors the chart-title bold path. legendRPr is RunProperties|DefaultRunProperties;
            // read the "b" attribute generically.
            var legendBold = legendRPr?.GetAttributes().FirstOrDefault(a => a.LocalName == "b").Value;
            info.LegendFontBold = legendBold == "1" || legendBold == "true";
            // #7f: honor <c:legendPos w:val="r|l|t|b|tr"/>.
            var posEl = legendEl.Elements().FirstOrDefault(e => e.LocalName == "legendPos");
            var posVal = posEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (!string.IsNullOrEmpty(posVal)) info.LegendPos = posVal!;
            // <c:legendEntry><c:idx val="N"/><c:delete val="1"/></c:legendEntry>
            // hides the legend swatch+label for series N (idx = series index for
            // bar/column/line/area). The series still plots; only its legend
            // entry is suppressed. No entries → empty set → all series shown.
            foreach (var entryEl in legendEl.Elements().Where(e => e.LocalName == "legendEntry"))
            {
                var entryDelEl = entryEl.Elements().FirstOrDefault(e => e.LocalName == "delete");
                var entryDelVal = entryDelEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                if (entryDelVal != "1" && entryDelVal != "true") continue;
                var idxEl = entryEl.Elements().FirstOrDefault(e => e.LocalName == "idx");
                var idxVal = idxEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                if (int.TryParse(idxVal, out var idx)) info.DeletedLegendEntries.Add(idx);
            }
        }
        else
        {
            // No <c:legend> element → PowerPoint/Excel render NO legend. Real
            // Office keys legend visibility strictly off the element's presence,
            // not off a series-count heuristic (verified against Office:
            // legend=none charts show no legend even with 2+ series). Guessing
            // here made legend=none still draw the Alpha/Beta swatches.
            info.HasLegend = false;
        }

        // Marker shapes, smooth, and dash per series
        if (chartTypeEl != null)
        {
            // Chart-level smooth (lineChart > smooth val="1")
            var chartSmooth = chartTypeEl.Elements().FirstOrDefault(e => e.LocalName == "smooth");
            var chartSmoothVal = chartSmooth?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            // CT_Boolean defaults to true: a bare <c:smooth/> (no val attr) means ON.
            // PowerPoint emits the bare form; the old `== "1"` read it as straight lines.
            var chartIsSmooth = chartSmooth != null
                && (string.IsNullOrEmpty(chartSmoothVal)
                    || (chartSmoothVal != "0" && !chartSmoothVal.Equals("false", StringComparison.OrdinalIgnoreCase)));

            // PowerPoint's <c:lineChart>/<c:scatterChart> emit a chart-level
            // <c:marker val="1"/> after all <c:ser> to opt every series into
            // the default marker cycle. Series without their own <c:marker>
            // get a shape chosen by series index (PowerPoint's built-in
            // sequence). Without the chart-level flag, unmarked series stay
            // marker-free. Matches real PowerPoint rendering of
            // /tmp/r14_v2.pptx chart3 (series A circle = explicit,
            // series B square = default cycle index 1).
            var chartLevelMarker = chartTypeEl.Elements()
                .Where(e => e.LocalName == "marker")
                .LastOrDefault(); // chart-level <c:marker val="1"/> appears after series
            var chartMarkerVal = chartLevelMarker?.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "val").Value;
            // CT_Boolean defaults to true: a bare <c:marker/> (no val attr) opts every series
            // into the default marker cycle. The old `== "1"` read it as markers-off.
            var chartMarkersOn = chartLevelMarker != null && (chartMarkerVal is null or "" or "1" or "true");
            // <c:scatterChart> uses <c:scatterStyle val="..."/> instead of a
            // chart-level <c:marker>. Values containing "marker" (lineMarker /
            // marker / smoothMarker) mean every series gets the default cycle.
            var scatterStyleEl = chartTypeEl.Elements().FirstOrDefault(e => e.LocalName == "scatterStyle");
            var scatterStyleVal = scatterStyleEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (scatterStyleVal != null && scatterStyleVal.Contains("arker", StringComparison.OrdinalIgnoreCase))
                chartMarkersOn = true;
            // Scatter charts carry their smooth-curve signal in <c:scatterStyle>
            // ("smooth"/"smoothMarker"), not in <c:smooth>. PowerPoint draws Bézier
            // curves for these; the renderer previously read only <c:smooth> and so
            // drew straight segments. Propagate it into the per-series smooth default.
            if (scatterStyleVal != null && scatterStyleVal.StartsWith("smooth", StringComparison.OrdinalIgnoreCase))
                chartIsSmooth = true;
            // R16c: scatterStyle exactly "marker" (or "none") = markers without a
            // connecting line. "lineMarker"/"smoothMarker"/"line"/"smooth" keep the
            // line. Suppress the polyline in that case.
            if (scatterStyleVal != null
                && (scatterStyleVal.Equals("marker", StringComparison.OrdinalIgnoreCase)
                    || scatterStyleVal.Equals("none", StringComparison.OrdinalIgnoreCase)))
                info.ScatterMarkersOnly = true;
            // Default cycle observed in PowerPoint line/scatter charts.
            var defaultMarkerCycle = new[] { "circle", "square", "diamond", "triangle", "x", "star", "plus", "dash", "dot" };

            int serIdx = 0;
            foreach (var ser in serElements)
            {
                var marker = ser.Elements().FirstOrDefault(e => e.LocalName == "marker");
                var symbol = marker?.Elements().FirstOrDefault(e => e.LocalName == "symbol");
                var symbolVal = symbol?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                string resolvedShape;
                if (symbolVal != null)
                    resolvedShape = symbolVal;
                else if (chartMarkersOn)
                    resolvedShape = defaultMarkerCycle[serIdx % defaultMarkerCycle.Length];
                else
                    resolvedShape = "none";
                info.MarkerShapes.Add(resolvedShape);
                var sizeEl = marker?.Elements().FirstOrDefault(e => e.LocalName == "size");
                var sizeVal = sizeEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                info.MarkerSizes.Add(sizeVal != null && int.TryParse(sizeVal, out var ms) ? ms : 5);
                // Marker fill + border (<c:marker><c:spPr>). PowerPoint paints a
                // marker with an explicit fill AND a series-colored outline; we
                // read the fill from solidFill and the border from <a:ln>. null
                // (no spPr) defers to the series color at the call site.
                var markerSpPr = marker?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                var mFill = ExtractFillColor(markerSpPr, themeColors);
                info.MarkerFillColors.Add(mFill != null ? $"#{mFill}" : null);
                var mLine = ExtractLineColor(markerSpPr, themeColors);
                info.MarkerLineColors.Add(mLine != null ? $"#{mLine}" : null);
                serIdx++;

                // Per-series smooth (overrides chart-level)
                var serSmooth = ser.Elements().FirstOrDefault(e => e.LocalName == "smooth");
                var serSmoothVal = serSmooth?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                info.Smooth.Add(serSmooth != null
                    ? (string.IsNullOrEmpty(serSmoothVal)
                        || (serSmoothVal != "0" && !serSmoothVal.Equals("false", StringComparison.OrdinalIgnoreCase)))
                    : chartIsSmooth);

                // Per-series dash pattern and line width
                var spPr = ser.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                var ln = spPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
                var prstDash = ln?.Elements().FirstOrDefault(e => e.LocalName == "prstDash");
                var dashVal = prstDash?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                info.LineDashes.Add(dashVal ?? "solid");

                // Per-series line width (a:ln w="..." in EMU, convert to pt: 1pt = 12700 EMU)
                var lnWidth = ln?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value;
                info.LineWidths.Add(lnWidth != null && int.TryParse(lnWidth, out var lw) ? Math.Round(lw / EmuConverter.EmuPerPointF, 1) : 2);

                // Per-series "no line" (a:ln/a:noFill): PowerPoint hides the connecting
                // polyline (markers only). The renderer always drew the polyline, ignoring it.
                info.SeriesLineHide.Add(ln?.Elements().Any(e => e.LocalName == "noFill") == true);

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
                    tlInfo.Color = ExtractLineColor(tlSpPr, themeColors);
                    if (tlLn?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value is string tlw
                        && int.TryParse(tlw, out var tlwPt))
                        tlInfo.Width = Math.Round(tlwPt / EmuConverter.EmuPerPointF, 1);
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
                    var ebNoCap = errBarsEl.Elements().FirstOrDefault(e => e.LocalName == "noEndCap");
                    ebInfo.NoEndCap = ebNoCap?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value is not ("0" or "false");
                    if (ebNoCap == null) ebInfo.NoEndCap = false;
                    // Read error value from Plus/Minus > NumLit > NumericPoint > v
                    var plusEl = errBarsEl.Elements().FirstOrDefault(e => e.LocalName == "plus");
                    var numPt = plusEl?.Descendants().FirstOrDefault(e => e.LocalName == "v");
                    if (numPt != null && double.TryParse(numPt.InnerText,
                        System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var ebVal))
                        ebInfo.Value = ebVal;
                    // Error bar styling
                    var ebSpPr = errBarsEl.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                    ebInfo.Color = ExtractLineColor(ebSpPr, themeColors);
                    var ebLn = ebSpPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
                    if (ebLn?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value is string ebw
                        && int.TryParse(ebw, out var ebwPt))
                        ebInfo.Width = Math.Round(ebwPt / EmuConverter.EmuPerPointF, 1);
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
                info.DropLineColor = ExtractLineColor(dlSpPr, themeColors);
                if (dlLn?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value is string dlw
                    && int.TryParse(dlw, out var dlwPt))
                    info.DropLineWidth = Math.Round(dlwPt / EmuConverter.EmuPerPointF, 1);
                var dlDash = dlLn?.Elements().FirstOrDefault(e => e.LocalName == "prstDash");
                info.DropLineDash = dlDash?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            }
            var hiLowEl = chartTypeEl.Elements().FirstOrDefault(e => e.LocalName == "hiLowLines");
            info.HasHighLowLines = hiLowEl != null;
            if (hiLowEl != null)
            {
                var hlSpPr = hiLowEl.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                var hlLn = hlSpPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
                info.HighLowLineColor = ExtractLineColor(hlSpPr, themeColors);
                if (hlLn?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value is string hlw
                    && int.TryParse(hlw, out var hlwPt))
                    info.HighLowLineWidth = Math.Round(hlwPt / EmuConverter.EmuPerPointF, 1);
            }
            var upDownBars = chartTypeEl.Elements().FirstOrDefault(e => e.LocalName == "upDownBars");
            info.HasUpDownBars = upDownBars != null;
            if (upDownBars != null)
            {
                var upSpPr = upDownBars.Elements().FirstOrDefault(e => e.LocalName == "upBars")
                    ?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                var dnSpPr = upDownBars.Elements().FirstOrDefault(e => e.LocalName == "downBars")
                    ?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                // Leave null when the up/down bars carry no fill: each renderer applies its
                // own default (line chart → green/red; stock candlestick → white/black per the
                // OOXML spec). ExtractFillColor resolves srgbClr/schemeClr/sysClr/prstClr.
                info.UpBarColor = ExtractFillColor(upSpPr, themeColors);
                info.DownBarColor = ExtractFillColor(dnSpPr, themeColors);
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

        // Reversed category axis (catAx maxMin): reverse the category order centrally so
        // every chart type (column/bar/line/area) renders the flipped order forward —
        // PowerPoint draws categories right-to-left (e.g. Mar,Feb,Jan). Reverse the
        // category labels, each series' values, and the per-point color overrides (keyed
        // by category index) in lockstep. Cat axis reversal doesn't apply to pie/doughnut
        // (no catAx orientation), and the horizontal-bar renderer's existing
        // first-category-at-bottom flip composes correctly (first category moves to top).
        if (info.IsCatReversed && info.Categories.Length > 1)
        {
            int n = info.Categories.Length;
            Array.Reverse(info.Categories);
            for (int i = 0; i < info.Series.Count; i++)
            {
                var v = info.Series[i].values;
                Array.Reverse(v);
                if (i < info.PerPointColors.Count && info.PerPointColors[i].Count > 0)
                    info.PerPointColors[i] = info.PerPointColors[i]
                        .ToDictionary(kv => n - 1 - kv.Key, kv => kv.Value);
            }
        }

        return info;
    }

    /// <summary>Extract series colors (per-point for pie/doughnut, stroke for line/scatter, fill for others).</summary>
    private static List<string> ExtractColors(List<OpenXmlElement> serElements, List<(string name, double[] values)> series,
        bool isPieType, string chartType, Dictionary<string, string>? themeColors = null, bool varyColors = true)
    {
        var colors = new List<string>();

        if (isPieType && serElements.Count > 0)
        {
            // Pie/doughnut: colors are per data point (dPt), not per series.
            var ser = serElements[0];
            var dPts = ser.Elements().Where(e => e.LocalName == "dPt").ToList();
            var catCount = series.FirstOrDefault().values?.Length ?? 0;
            // <c:varyColors val="0"/>: PowerPoint colors every non-overridden slice in
            // the SINGLE series color (monochrome pie) instead of cycling the accent
            // palette. Default (absent or val="1") varies by point. Explicit dPt fills
            // still win in both modes.
            string? serColorUniform = null;
            if (!varyColors)
            {
                var serSpPr = ser.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                var serRgb = ExtractFillColor(serSpPr, themeColors);
                serColorUniform = serRgb != null ? $"#{serRgb}" : FallbackColors[0];
            }
            for (int i = 0; i < catCount; i++)
            {
                var dPt = dPts.FirstOrDefault(d =>
                {
                    var idxEl = d.Elements().FirstOrDefault(e => e.LocalName == "idx");
                    if (idxEl == null) return false;
                    return idxEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == i.ToString();
                });
                var rgb = ExtractFillColor(dPt?.Elements().FirstOrDefault(e => e.LocalName == "spPr"), themeColors);
                colors.Add(rgb != null ? $"#{rgb}"
                    : serColorUniform ?? FallbackColors[i % FallbackColors.Length]);
            }
        }
        else
        {
            // Detect line/scatter series PER-SERIES from the owning chart group, not a
            // single chart-level flag: a combo chart mixes bar and line groups, so the
            // line series' color lives in <a:ln><a:solidFill> (stroke) while the bar
            // series' lives in <a:solidFill> (fill). A single chartType=="combo" flag
            // matched neither, so line series rendered fallback colors.
            for (int i = 0; i < series.Count; i++)
            {
                string? rgb = null;
                if (i < serElements.Count)
                {
                    var parentName = (serElements[i].Parent?.LocalName ?? "").ToLowerInvariant();
                    var serIsLine = parentName.Contains("line") || parentName.Contains("scatter");
                    var spPr = serElements[i].Elements().FirstOrDefault(e => e.LocalName == "spPr");
                    if (serIsLine)
                    {
                        // For line/scatter, prefer stroke color from a:ln > a:solidFill
                        var ln = spPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
                        rgb = ExtractFillColor(ln, themeColors);
                    }
                    // Fallback to solidFill
                    rgb ??= ExtractFillColor(spPr, themeColors);
                }
                colors.Add(rgb != null ? $"#{rgb}" : FallbackColors[i % FallbackColors.Length]);
            }
        }
        return colors;
    }

    /// <summary>Extract per-data-point fill overrides (<c:dPt>) for non-pie
    /// charts (bar/column). Returns one dict per series mapping zero-based
    /// category idx -> '#'-prefixed hex. Series with no dPt yield an empty dict,
    /// so the renderer falls back to the per-series color (regression-safe).
    /// Colors resolve through the same ExtractFillColor path used for series
    /// fills (srgbClr/schemeClr/theme).</summary>
    private static List<Dictionary<int, string>> ExtractPerPointColors(
        List<OpenXmlElement> serElements, Dictionary<string, string>? themeColors = null)
    {
        var result = new List<Dictionary<int, string>>();
        foreach (var ser in serElements)
        {
            var map = new Dictionary<int, string>();
            foreach (var dPt in ser.Elements().Where(e => e.LocalName == "dPt"))
            {
                var idxEl = dPt.Elements().FirstOrDefault(e => e.LocalName == "idx");
                var idxStr = idxEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                if (!int.TryParse(idxStr, out var idx)) continue;
                var rgb = ExtractFillColor(dPt.Elements().FirstOrDefault(e => e.LocalName == "spPr"), themeColors);
                if (rgb != null) map[idx] = $"#{rgb}";
            }
            result.Add(map);
        }
        return result;
    }

    /// <summary>Per-series set of data-point indices whose data label was explicitly
    /// deleted (&lt;c:dLbls&gt;&lt;c:dLbl&gt;&lt;c:idx&gt;&lt;c:delete/&gt;). PowerPoint hides
    /// just those points' labels while keeping the rest of the series' labels.</summary>
    private static List<HashSet<int>> ExtractDeletedLabels(List<OpenXmlElement> serElements)
    {
        var result = new List<HashSet<int>>();
        foreach (var ser in serElements)
        {
            var set = new HashSet<int>();
            var dLbls = ser.Elements().FirstOrDefault(e => e.LocalName == "dLbls");
            if (dLbls != null)
                foreach (var dLbl in dLbls.Elements().Where(e => e.LocalName == "dLbl"))
                {
                    var del = dLbl.Elements().FirstOrDefault(e => e.LocalName == "delete");
                    if (del == null) continue;
                    var dv = del.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                    if (dv is not (null or "" or "1" or "true")) continue; // CT_Boolean default true
                    var idxStr = dLbl.Elements().FirstOrDefault(e => e.LocalName == "idx")
                        ?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                    if (int.TryParse(idxStr, out var idx)) set.Add(idx);
                }
            result.Add(set);
        }
        return result;
    }

    /// <summary>Extract per-series <c:invertIfNegative>. PowerPoint's effective
    /// default is TRUE (negative bars render hollow) when the element is absent;
    /// FALSE only when explicitly val="0". Returns one bool per series, aligned
    /// to series order.</summary>
    private static List<bool> ExtractInvertIfNegative(List<OpenXmlElement> serElements, int seriesCount)
    {
        var result = new List<bool>();
        for (int i = 0; i < seriesCount; i++)
        {
            // Default true when the element is absent. When present, honor its
            // val (val="0"/false → keep negatives solid; val="1"/absent-attr → true).
            var invEl = i < serElements.Count
                ? serElements[i].Elements().FirstOrDefault(e => e.LocalName == "invertIfNegative")
                : null;
            if (invEl == null) { result.Add(true); continue; }
            var valStr = invEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            // An <c:invertIfNegative/> with no val attribute defaults to true.
            result.Add(string.IsNullOrEmpty(valStr) || (valStr != "0" && !valStr.Equals("false", StringComparison.OrdinalIgnoreCase)));
        }
        return result;
    }

    /// <summary>Extract per-series fill opacity from the series spPr
    /// (pie/doughnut: per data-point dPt spPr) <a:solidFill><a:alpha val="…"/>.
    /// Returns null per entry when no explicit alpha is declared, so the
    /// renderer uses its opaque (1.0) default for that series.</summary>
    private static List<double?> ExtractFillOpacities(List<OpenXmlElement> serElements,
        List<(string name, double[] values)> series, bool isPieType)
    {
        var opacities = new List<double?>();
        if (isPieType && serElements.Count > 0)
        {
            var ser = serElements[0];
            var dPts = ser.Elements().Where(e => e.LocalName == "dPt").ToList();
            var catCount = series.FirstOrDefault().values?.Length ?? 0;
            for (int i = 0; i < catCount; i++)
            {
                var dPt = dPts.FirstOrDefault(d =>
                    d.Elements().FirstOrDefault(e => e.LocalName == "idx")
                        ?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == i.ToString());
                opacities.Add(ExtractFillAlpha(dPt?.Elements().FirstOrDefault(e => e.LocalName == "spPr")));
            }
        }
        else
        {
            for (int i = 0; i < series.Count; i++)
                opacities.Add(i < serElements.Count
                    ? ExtractFillAlpha(serElements[i].Elements().FirstOrDefault(e => e.LocalName == "spPr"))
                    : null);
        }
        return opacities;
    }

    /// <summary>Extract the alpha (0..1) from solidFill > srgbClr/schemeClr >
    /// a:alpha (val is /100000) inside an spPr. Null when absent.</summary>
    private static double? ExtractFillAlpha(OpenXmlElement? spPr)
    {
        if (spPr == null) return null;
        var solidFill = spPr.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        var clr = solidFill?.Elements().FirstOrDefault(e => e.LocalName is "srgbClr" or "schemeClr");
        var alphaEl = clr?.Elements().FirstOrDefault(e => e.LocalName == "alpha");
        if (alphaEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value is string av
            && int.TryParse(av, out var a) && a >= 0 && a <= 100000)
            return a / 100000.0;
        return null;
    }

    /// <summary>Extract hex color (without #) from solidFill > srgbClr (or schemeClr
    /// resolved against the theme map) inside an spPr or ln element.</summary>
    private static string? ExtractFillColor(OpenXmlElement? container, Dictionary<string, string>? themeColors = null)
    {
        if (container == null) return null;
        var solidFill = container.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        var srgb = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        var v = srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        // schemeClr (e.g. accent3): resolve through the theme color map so a
        // series styled with a scheme color renders its actual theme hex instead
        // of dropping to the wrong fallback-palette index. Mirrors the shape
        // renderer's ResolveFillColor (schemeClr → themeColors[name] → hex).
        if (v == null && themeColors != null)
        {
            var scheme = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
            var schemeName = scheme?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (!string.IsNullOrEmpty(schemeName))
            {
                var canonical = ParseHelpers.NormalizeSchemeColorName(schemeName) ?? schemeName;
                if (themeColors.TryGetValue(canonical, out var themeHex)
                    || themeColors.TryGetValue(schemeName, out themeHex))
                    v = themeHex;
            }
        }
        // gradFill fallback: a gradient-filled series has no solidFill. SVG bar
        // fills are flat, so approximate the gradient with its FIRST stop color
        // (the gradient start) rather than dropping to the wrong fallback accent.
        if (v == null)
        {
            var gradFill = container.Elements().FirstOrDefault(e => e.LocalName == "gradFill");
            var gsLst = gradFill?.Elements().FirstOrDefault(e => e.LocalName == "gsLst");
            var firstGs = gsLst?.Elements().FirstOrDefault(e => e.LocalName == "gs");
            var gsSrgb = firstGs?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
            v = gsSrgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            // First stop may be a schemeClr (e.g. accent2): resolve through the theme map,
            // mirroring the solidFill schemeClr branch above. Without this a gradient series
            // whose first stop is a theme color dropped to the wrong fallback-palette accent.
            if (v == null && themeColors != null && firstGs != null)
            {
                var gsScheme = firstGs.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
                var gsName = gsScheme?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                if (!string.IsNullOrEmpty(gsName))
                {
                    var canonical = ParseHelpers.NormalizeSchemeColorName(gsName) ?? gsName;
                    if (themeColors.TryGetValue(canonical, out var themeHex)
                        || themeColors.TryGetValue(gsName, out themeHex))
                        v = themeHex;
                }
            }
        }
        // pattFill fallback: a pattern-filled series (Format Data Series -> Pattern Fill)
        // has no solidFill. SVG bar fills are flat, so approximate with the pattern's
        // FOREGROUND color (same flat-approximation as the gradFill case above) instead of
        // dropping to the wrong fallback accent. The stripe texture itself is not rendered
        // (would need SVG <pattern> defs) — surfacing the fg color is the consistent
        // approximation. Resolve fg srgbClr, else schemeClr via the theme map.
        if (v == null)
        {
            var pattFill = container.Elements().FirstOrDefault(e => e.LocalName == "pattFill");
            var fgClr = pattFill?.Elements().FirstOrDefault(e => e.LocalName == "fgClr");
            v = fgClr?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr")?
                .GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (v == null && fgClr != null && themeColors != null)
            {
                var sName = fgClr.Elements().FirstOrDefault(e => e.LocalName == "schemeClr")?
                    .GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                if (!string.IsNullOrEmpty(sName))
                {
                    var canon = ParseHelpers.NormalizeSchemeColorName(sName) ?? sName;
                    if (themeColors.TryGetValue(canon, out var hx) || themeColors.TryGetValue(sName, out hx))
                        v = hx;
                }
            }
        }
        // Reject non-hex values — the return flows into $"#{...}" inline SVG
        // fill/style attributes. Same XSS class as w:color / w:shd / border.
        if (v == null) return null;
        if (v.Length is not (3 or 6 or 8)) return null;
        foreach (var c in v)
            if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'))) return null;
        return v;
    }

    /// <summary>Extract font color from RunProperties or DefaultRunProperties
    /// (solidFill > srgbClr, or solidFill > schemeClr resolved through the theme).</summary>
    private static string? ExtractFontColor(OpenXmlElement? rPr, Dictionary<string, string>? themeColors = null)
    {
        if (rPr == null) return null;
        var solidFill = rPr.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        var srgb = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        var val = srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        // schemeClr (accent1.., tx1, bg1, ...): resolve through the theme color map
        // so a chart title / axis / legend styled with a scheme color renders its
        // actual theme hex instead of dropping to the global default text color.
        // Mirrors ExtractFillColor's schemeClr branch.
        if (val == null && themeColors != null)
        {
            var scheme = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
            var schemeName = scheme?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (!string.IsNullOrEmpty(schemeName))
            {
                var canonical = ParseHelpers.NormalizeSchemeColorName(schemeName) ?? schemeName;
                if (themeColors.TryGetValue(canonical, out var themeHex)
                    || themeColors.TryGetValue(schemeName, out themeHex))
                    val = themeHex;
            }
        }
        return HexOrNull(val);
    }

    /// <summary>Read an axis tick-label rotation (in degrees) from its
    /// &lt;c:txPr&gt;&lt;a:bodyPr rot="..."/&gt;. OOXML rot is 1/60000 degree.
    /// Returns null when txPr / bodyPr / rot is absent or rot is 0 (so plain
    /// charts keep horizontal labels, regression-safe).</summary>
    private static int? ExtractAxisLabelRotationDeg(OpenXmlElement? txPr)
    {
        if (txPr == null) return null;
        var bodyPr = txPr.Elements().FirstOrDefault(e => e.LocalName == "bodyPr");
        var rotVal = bodyPr?.GetAttributes().FirstOrDefault(a => a.LocalName == "rot").Value;
        if (rotVal == null || !int.TryParse(rotVal, out var rot) || rot == 0) return null;
        return rot / 60000;
    }

    /// <summary>Extract line/outline color from spPr (ln > solidFill > srgbClr, or
    /// > schemeClr resolved through the theme).</summary>
    private static string? ExtractLineColor(OpenXmlElement? spPr, Dictionary<string, string>? themeColors = null)
    {
        if (spPr == null) return null;
        var ln = spPr.Elements().FirstOrDefault(e => e.LocalName == "ln");
        if (ln == null) return null;
        var solidFill = ln.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        var srgb = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        var val = srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        // schemeClr (accent1.., tx1, bg1, ...): resolve through the theme map so a
        // gridline / axis line / trendline / error bar / drop line / hi-low line /
        // marker line / plot or chart border styled with a scheme color renders its
        // theme hex instead of falling back to a default or the series color.
        // Mirrors ExtractFillColor / ExtractFontColor.
        if (val == null && themeColors != null)
        {
            var scheme = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
            var schemeName = scheme?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (!string.IsNullOrEmpty(schemeName))
            {
                var canonical = ParseHelpers.NormalizeSchemeColorName(schemeName) ?? schemeName;
                if (themeColors.TryGetValue(canonical, out var themeHex)
                    || themeColors.TryGetValue(schemeName, out themeHex))
                    val = themeHex;
            }
        }
        return HexOrNull(val);
    }

    /// <summary>Read a:ln/@w (EMU) from an spPr. Null when no a:ln or no width
    /// attribute (caller defaults to ~0.75pt for a present-but-widthless line).</summary>
    private static long? ExtractLineWidthEmu(OpenXmlElement? spPr)
    {
        var ln = spPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
        var w = ln?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value;
        return long.TryParse(w, out var emu) ? emu : (long?)null;
    }

    /// <summary>EMU outline width → SVG stroke px (1 EMU = 1/914400 in, pt = EMU/12700,
    /// px ≈ pt * 4/3). Null width = PowerPoint's default ~0.75pt line.</summary>
    private static double EmuToStrokePx(long? emu)
    {
        var pt = emu.HasValue ? emu.Value / 12700.0 : 0.75;
        return pt * 4.0 / 3.0;
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

    /// <summary>True when a <c:majorTickMark val="..."/> value should draw ticks
    /// (present and not "none"). PowerPoint draws short perpendicular lines at
    /// each major label for "out"/"in"/"cross". Absent/"none" => no ticks.</summary>
    private static bool TickMarkVisible(string? v)
        => v != null && v != "none";

    /// <summary>Emit a single major tick mark on a vertical axis (value axis on the
    /// left, or horizontal-bar category axis): a short horizontal line at y.
    /// "out" extends left of the axis (x=axisX), "in" right, "cross" straddles.</summary>
    private void EmitVAxisTick(StringBuilder sb, double axisX, double y, string mode)
    {
        double x1 = axisX, x2 = axisX;
        if (mode == "out") { x1 = axisX - MajorTickLen; x2 = axisX; }
        else if (mode == "in") { x1 = axisX; x2 = axisX + MajorTickLen; }
        else if (mode == "cross") { x1 = axisX - MajorTickLen; x2 = axisX + MajorTickLen; }
        sb.AppendLine($"        <line x1=\"{x1:0.#}\" y1=\"{y:0.#}\" x2=\"{x2:0.#}\" y2=\"{y:0.#}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
    }

    /// <summary>Emit a single major tick mark on a horizontal axis (category axis at
    /// bottom, or horizontal-bar value axis): a short vertical line at x.
    /// "out" extends below the axis (y=axisY), "in" above, "cross" straddles.</summary>
    private void EmitHAxisTick(StringBuilder sb, double x, double axisY, string mode)
    {
        double y1 = axisY, y2 = axisY;
        if (mode == "out") { y1 = axisY; y2 = axisY + MajorTickLen; }
        else if (mode == "in") { y1 = axisY - MajorTickLen; y2 = axisY; }
        else if (mode == "cross") { y1 = axisY - MajorTickLen; y2 = axisY + MajorTickLen; }
        sb.AppendLine($"        <line x1=\"{x:0.#}\" y1=\"{y1:0.#}\" x2=\"{x:0.#}\" y2=\"{y2:0.#}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
    }

    /// <summary>Normalize a chart color for direct emission into an SVG
    /// fill/stroke attribute or CSS color. Bare OOXML hex ("FF0000") gets a
    /// '#' prefix; values already '#'-prefixed or non-hex (named/scheme
    /// colors handled upstream) pass through unchanged so we never double the
    /// '#'. Mirrors the $"#{rgb}" pattern used by ExtractColors.</summary>
    internal static string CssHexColor(string v)
        => HexOrNull(v) != null ? "#" + v : v;

    /// <summary>Render the chart SVG content (inside an already-opened svg tag) based on ChartInfo.</summary>
    public void RenderChartSvgContent(StringBuilder sb, ChartInfo info, int svgW, int svgH,
        int marginLeft = 45, int marginTop = 10, int marginRight = 15, int marginBottom = 30)
    {
        // Sync instance font sizes and colors from ChartInfo
        ValFontPx = info.ValFontPx;
        CatFontPx = info.CatFontPx;
        // These ChartInfo colors are stored as raw OOXML hex (no '#'); the SVG
        // fill/stroke attributes need '#'-prefixed CSS hex or browsers render
        // the element black. Route every one through CssHexColor so a bare
        // "FF0000" becomes "#FF0000" while named/already-#'d values pass through.
        if (info.ValFontColor != null) AxisColor = CssHexColor(info.ValFontColor);
        if (info.CatFontColor != null) CatColor = CssHexColor(info.CatFontColor);
        ValTickLabelsBold = info.ValFontBold;
        CatTickLabelsBold = info.CatFontBold;
        ValTickLabelsItalic = info.ValFontItalic;
        CatTickLabelsItalic = info.CatFontItalic;
        if (info.GridlineColor != null) GridColor = CssHexColor(info.GridlineColor);
        GridlineDash = info.GridlineDash;
        GridlineWidthPx = info.GridlineWidthEmu.HasValue ? EmuToStrokePx(info.GridlineWidthEmu) : 0.5;
        ShowValGridlines = info.ValMajorGridlines;
        ShowCatGridlines = info.CatMajorGridlines;
        ShowValMinorGridlines = info.ValMinorGridlines;
        ShowCatMinorGridlines = info.CatMinorGridlines;
        ValAxisVisible = info.ValAxisVisible;
        CatAxisVisible = info.CatAxisVisible;
        ValTickLabelsHidden = info.ValTickLabelsHidden;
        CatTickLabelsHidden = info.CatTickLabelsHidden;
        ValMajorTickMark = info.ValMajorTickMark;
        CatMajorTickMark = info.CatMajorTickMark;
        CatTickLabelSkip = info.CatTickLabelSkip;
        PerPointDeletedLabels = info.PerPointDeletedLabels;
        if (info.AxisLineColor != null) AxisLineColor = CssHexColor(info.AxisLineColor);
        DataLabelFontPx = info.DataLabelFontPx;
        // Minor-gridline count from <c:minorUnit> / <c:majorUnit> (sub-intervals
        // per major band). PowerPoint draws (majorUnit/minorUnit - 1) lines per
        // band; the renderer's loops divide the band into MinorGridlineCount parts.
        if (info.MinorUnit is > 0 && info.MajorUnit is > 0)
            MinorGridlineCount = Math.Max(1, (int)Math.Round(info.MajorUnit.Value / info.MinorUnit.Value));
        if (info.DataLabelFontColor != null) DataLabelColor = CssHexColor(info.DataLabelFontColor);
        DataLabelBold = info.DataLabelBold;
        DataLabelItalic = info.DataLabelItalic;
        DataLabelPos = info.DataLabelPos;
        HasExplicitDataLabelPos = info.HasExplicitDataLabelPos;
        FirstSliceAngle = info.FirstSliceAngle;
        SeriesFillOpacities = info.SeriesFillOpacities;
        InvertIfNegative = info.InvertIfNegative;
        ValAxisUnitDivisor = info.ValueAxisUnitDivisor;

        // Increase right margin for long axis labels (e.g. "$1,000,000")
        if (!string.IsNullOrEmpty(info.ValNumFmt) && marginRight < 30)
            marginRight = 30;

        // Rotated category labels (catAx txPr bodyPr rot) hang diagonally below
        // the axis and trail toward the side, so reserve extra bottom (and, for
        // the leading label, left) space so they aren't clipped. Approximate the
        // longest label's pixel length, then project it onto the rotation angle.
        if (info.CatAxisLabelRotationDeg is int catRot && catRot != 0
            && info.Categories.Length > 0)
        {
            var maxLen = info.Categories.Max(c => (c ?? "").Length);
            var labelPx = maxLen * info.CatFontPx * 0.5;
            var rad = Math.Abs(catRot) * Math.PI / 180.0;
            var extraBottom = (int)(labelPx * Math.Sin(rad)) + 4;
            marginBottom += extraBottom;
            var extraSide = (int)(labelPx * Math.Cos(rad)) + 4;
            if (catRot < 0 && marginLeft < extraSide) marginLeft = extraSide;
            else if (catRot > 0 && marginRight < extraSide) marginRight = extraSide;
        }

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

        if (chartType == "pieOfPie" || chartType == "barOfPie")
        {
            // R16-1: pieOfPie / barOfPie render a main pie PLUS a secondary plot
            // (a small pie or a bar cluster) for the trailing data points, joined
            // by connector lines. Must branch before the generic Contains("pie")
            // test, which would otherwise render a plain single pie.
            RenderOfPieChartSvg(sb, info.Series, info.Categories, info.Colors,
                marginLeft, marginTop, plotW, plotH, chartType == "barOfPie",
                info.ShowDataLabels, info.ShowDataLabelVal, info.ShowDataLabelPercent,
                info.DataLabelsNumFmt);
        }
        else if (chartType.Contains("pie") || chartType.Contains("doughnut"))
        {
            if (info.Is3D)
                RenderPie3DSvg(sb, info.Series, info.Categories, info.Colors, svgW, svgH,
                    info.ShowDataLabels, info.ShowDataLabelVal, info.ShowDataLabelPercent,
                    info.RotateX > 0 ? info.RotateX : 30);
            else
                RenderPieChartSvg(sb, info.Series, info.Categories, info.Colors, svgW, svgH, info.HoleRatio, info.ShowDataLabels,
                    info.ShowDataLabelVal, info.ShowDataLabelPercent, info.ShowDataLabelCatName, info.Explosions,
                    info.ShowDataLabelSerName);
        }
        else if (chartType.Contains("area"))
        {
            var areaW = plotW - (int)(plotW * 0.03);
            if (info.Is3D)
                RenderArea3DSvg(sb, info.Series, info.Categories, info.Colors, marginLeft, marginTop, areaW, plotH,
                    info.IsStacked, info.RotateX, info.RotateY);
            else
                RenderAreaChartSvg(sb, info.Series, info.Categories, info.Colors, marginLeft, marginTop, areaW, plotH, info.IsStacked, info.IsPercent,
                    info.AxisMin, info.AxisMax, info.MajorUnit, info.ValNumFmt,
                    info.ShowDataLabels, info.ShowDataLabelVal, info.ShowDataLabelSerName,
                    info.ShowDataLabelCatName, info.DataLabelsNumFmt,
                    info.CatAxisLabelRotationDeg, info.ValAxisLabelRotationDeg,
                    info.IsReversed, info.Trendlines);
        }
        else if (chartType == "combo")
        {
            RenderComboChartSvg(sb, info.PlotArea!, info.Series, info.Categories, info.Colors, marginLeft, marginTop, plotW, plotH,
                info.ShowDataLabels, info.DataLabelsNumFmt, info.AxisMax);
        }
        else if (chartType.Contains("radar"))
        {
            RenderRadarChartSvg(sb, info.Series, info.Categories, info.Colors, svgW, svgH, CatFontPx, info.RadarStyle,
                info.ShowDataLabels, info.ShowDataLabelVal, info.ShowDataLabelSerName,
                info.ShowDataLabelCatName, info.DataLabelsNumFmt);
        }
        else if (chartType == "bubble")
        {
            RenderBubbleChartSvg(sb, info.PlotArea!, info.Series, info.Categories, info.Colors, marginLeft, marginTop, plotW, plotH,
                info.ShowDataLabels, info.ShowDataLabelVal, info.ShowDataLabelSerName,
                info.ShowDataLabelCatName, info.DataLabelsNumFmt);
        }
        else if (chartType == "stock")
        {
            RenderStockChartSvg(sb, info.PlotArea!, info.Series, info.Categories, info.Colors, marginLeft, marginTop, plotW, plotH, info.UpBarColor, info.DownBarColor);
        }
        else if (chartType == "scatter" && info.PlotArea != null)
        {
            // Scatter is an XY chart: both axes are numeric value axes. Route to
            // the dedicated renderer so X tick labels reflect the real X domain
            // (xVal) instead of the line renderer's 0..n category index.
            RenderScatterChartSvg(sb, info.PlotArea, info.Series, info.Colors, marginLeft, marginTop, plotW, plotH,
                info.MarkerShapes, info.MarkerSizes, info.LineWidths, info.LineDashes,
                info.ScatterMarkersOnly, info.ShowDataLabels,
                info.AxisMin, info.AxisMax, info.MajorUnit, info.ValNumFmt,
                info.Smooth, info.Trendlines, info.ErrorBars,
                info.MarkerFillColors, info.MarkerLineColors,
                info.DataLabelsNumFmt,
                info.ShowDataLabelVal, info.ShowDataLabelSerName, info.ShowDataLabelCatName,
                info.SeriesLineHide);
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
                    info.Trendlines, info.ErrorBars, info.ScatterMarkersOnly,
                    info.IsStacked, info.IsPercent, info.DataLabelsNumFmt,
                    info.MarkerFillColors, info.MarkerLineColors,
                    info.ShowDataLabelSerName, info.ShowDataLabelCatName,
                    info.ShowDataLabelVal || info.ShowDataLabelPercent,
                    info.CatAxisLabelRotationDeg, info.ValAxisLabelRotationDeg,
                    info.SeriesLineHide);
        }
        else
        {
            // Column/bar variants
            var isHorizontal = chartType.Contains("bar") && !chartType.Contains("column");
            // Structural signal that <c:overlap> was read and applied to the
            // clustered bar geometry. Emitted only when the element is present
            // in the chart XML (info.Overlap.HasValue) so a default chart does
            // not gain a spurious attribute. The geometry change lives in
            // RenderBarChartSvg; this is the inspectable marker.
            if (info.Overlap.HasValue)
                sb.AppendLine($"    <g data-overlap=\"{info.Overlap.Value}\"></g>");
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
                    info.IsWaterfall, info.ErrorBars,
                    info.IsPercent && info.ShowDataLabelPercent && !info.ShowDataLabelVal,
                    info.DataLabelsNumFmt, info.Overlap, info.IsReversed, info.PerPointColors,
                    info.CatAxisLabelRotationDeg, info.ValAxisLabelRotationDeg, info.Trendlines,
                    info.ShowDataLabelSerName, info.ShowDataLabelCatName,
                    info.ShowDataLabelVal || info.ShowDataLabelPercent, info.LogBase);
        }

        // Plot-area border (<c:plotArea><c:spPr><a:ln>). Drawn AFTER the plot
        // fill, gridlines, and series so the outline sits on top — matching how
        // PowerPoint traces the plot rectangle. No a:ln => no border (default).
        // Horizontal bar plots use a different geometry (handled inside
        // RenderBarChartSvg) so skip them here, same as the plot fill.
        if (info.PlotBorderColor != null && !isHorizBarType)
        {
            var pw = EmuToStrokePx(info.PlotBorderWidthEmu);
            sb.AppendLine($"    <rect x=\"{marginLeft}\" y=\"{marginTop}\" width=\"{plotW}\" height=\"{plotH}\" fill=\"none\" stroke=\"{CssHexColor(info.PlotBorderColor)}\" stroke-width=\"{pw:0.##}\"/>");
        }

        // Axis titles inside SVG — for horizontal bar charts, value axis is on bottom and category axis is on left
        var isHorizBar = chartType.Contains("bar") && !chartType.Contains("column");
        // Bubble/scatter have no category axis: the X axis is the primary value
        // axis and the Y axis is the secondary value axis.
        var isXY = chartType == "bubble" || chartType == "scatter";
        string? bottomTitle; int bottomTitleFont; bool bottomTitleBold; string? bottomTitleColor; bool bottomTitleItalic;
        string? leftTitle; int leftTitleFont; bool leftTitleBold; string? leftTitleColor; bool leftTitleItalic;
        if (isXY)
        {
            bottomTitle = info.ValAxisTitle; bottomTitleFont = info.ValAxisTitleFontPx; bottomTitleBold = info.ValAxisTitleBold; bottomTitleColor = info.ValAxisTitleColor; bottomTitleItalic = info.ValAxisTitleItalic;
            leftTitle = info.SecondaryValAxisTitle; leftTitleFont = info.SecondaryValAxisTitleFontPx; leftTitleBold = info.SecondaryValAxisTitleBold; leftTitleColor = info.SecondaryValAxisTitleColor; leftTitleItalic = info.SecondaryValAxisTitleItalic;
        }
        else
        {
            bottomTitle = isHorizBar ? info.ValAxisTitle : info.CatAxisTitle;
            bottomTitleFont = isHorizBar ? info.ValAxisTitleFontPx : info.CatAxisTitleFontPx;
            bottomTitleBold = isHorizBar ? info.ValAxisTitleBold : info.CatAxisTitleBold;
            bottomTitleColor = isHorizBar ? info.ValAxisTitleColor : info.CatAxisTitleColor;
            bottomTitleItalic = isHorizBar ? info.ValAxisTitleItalic : info.CatAxisTitleItalic;
            leftTitle = isHorizBar ? info.CatAxisTitle : info.ValAxisTitle;
            leftTitleFont = isHorizBar ? info.CatAxisTitleFontPx : info.ValAxisTitleFontPx;
            leftTitleBold = isHorizBar ? info.CatAxisTitleBold : info.ValAxisTitleBold;
            leftTitleColor = isHorizBar ? info.CatAxisTitleColor : info.ValAxisTitleColor;
            leftTitleItalic = isHorizBar ? info.CatAxisTitleItalic : info.ValAxisTitleItalic;
        }
        if (!string.IsNullOrEmpty(leftTitle))
            sb.AppendLine($"    <text x=\"10\" y=\"{svgH / 2}\" fill=\"{(leftTitleColor != null ? CssHexColor(leftTitleColor) : AxisColor)}\" font-size=\"{leftTitleFont}\"{(leftTitleBold ? " font-weight=\"bold\"" : "")}{(leftTitleItalic ? " font-style=\"italic\"" : "")} text-anchor=\"middle\" dominant-baseline=\"middle\" transform=\"rotate(-90,10,{svgH / 2})\">{HtmlEncode(leftTitle)}</text>");
        if (!string.IsNullOrEmpty(bottomTitle))
            sb.AppendLine($"    <text x=\"{svgW / 2}\" y=\"{svgH - 2}\" fill=\"{(bottomTitleColor != null ? CssHexColor(bottomTitleColor) : AxisColor)}\" font-size=\"{bottomTitleFont}\"{(bottomTitleBold ? " font-weight=\"bold\"" : "")}{(bottomTitleItalic ? " font-style=\"italic\"" : "")} text-anchor=\"middle\">{HtmlEncode(bottomTitle)}</text>");
        // Combo charts (non-XY): the secondary value axis title sits on the right.
        if (!isXY && !string.IsNullOrEmpty(info.SecondaryValAxisTitle))
            sb.AppendLine($"    <text x=\"{svgW - 10}\" y=\"{svgH / 2}\" fill=\"{(info.SecondaryValAxisTitleColor != null ? CssHexColor(info.SecondaryValAxisTitleColor) : SecondaryAxisColor)}\" font-size=\"{info.SecondaryValAxisTitleFontPx}\"{(info.SecondaryValAxisTitleBold ? " font-weight=\"bold\"" : "")}{(info.SecondaryValAxisTitleItalic ? " font-style=\"italic\"" : "")} text-anchor=\"middle\" dominant-baseline=\"middle\" transform=\"rotate(90,{svgW - 10},{svgH / 2})\">{HtmlEncode(info.SecondaryValAxisTitle)}</text>");

        // Display-units annotation (<c:dispUnits><c:dispUnitsLbl>, e.g. "Millions").
        // PowerPoint draws it beside the value axis. For a vertical value axis it
        // sits rotated -90° just inboard of the axis title; for a horizontal-bar /
        // XY value axis (which runs along the bottom) it sits centered below.
        if (!string.IsNullOrEmpty(info.ValueAxisUnitLabel))
        {
            if (isHorizBar || isXY)
            {
                var uy = svgH - (string.IsNullOrEmpty(bottomTitle) ? 4 : 14);
                sb.AppendLine($"    <text x=\"{svgW / 2}\" y=\"{uy}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\"{ValTickWeightAttr} text-anchor=\"middle\">{HtmlEncode(info.ValueAxisUnitLabel)}</text>");
            }
            else
            {
                var ux = string.IsNullOrEmpty(leftTitle) ? 12 : 26;
                sb.AppendLine($"    <text x=\"{ux}\" y=\"{svgH / 2}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\"{ValTickWeightAttr} text-anchor=\"middle\" dominant-baseline=\"middle\" transform=\"rotate(-90,{ux},{svgH / 2})\">{HtmlEncode(info.ValueAxisUnitLabel)}</text>");
            }
        }
    }

    /// <summary>Render chart legend HTML (outside the svg tag).</summary>
    public void RenderLegendHtml(StringBuilder sb, ChartInfo info, string fontColor = "#555")
    {
        if (!info.HasLegend) return;
        var legendColor = info.LegendFontColor != null ? CssHexColor(info.LegendFontColor) : fontColor;
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
        sb.Append($"<div class=\"chart-legend\" data-legend-pos=\"{safePos}\" style=\"{layoutCss};font-size:{info.LegendFontSize};color:{legendColor}{(info.LegendFontBold ? ";font-weight:bold" : "")}\">");
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
                // <c:legendEntry> delete: hide this series' swatch+label (it still plots).
                if (info.DeletedLegendEntries.Contains(i)) continue;
                var color = i < info.Colors.Count ? info.Colors[i] : DefaultColors[i % DefaultColors.Length];
                // Line/scatter legend keys are a short line segment (matching the
                // series stroke width + dash) plus the series marker, not a filled
                // square — PowerPoint keys the legend by chart type. The marker is
                // overlaid only when the series genuinely carries one (faithful to
                // OOXML; no default markers added).
                var isLineLegend = info.ChartType.Contains("line") || info.ChartType.Contains("scatter");
                string swatch;
                if (isLineLegend)
                {
                    var lw = i < info.LineWidths.Count && info.LineWidths[i] > 0 ? info.LineWidths[i] : 2.0;
                    var dash = i < info.LineDashes.Count ? info.LineDashes[i] : "solid";
                    var dashAttr = !string.IsNullOrEmpty(dash) && dash != "solid" ? $" stroke-dasharray=\"{RefLineDashArray(dash)}\"" : "";
                    var mShape = i < info.MarkerShapes.Count ? info.MarkerShapes[i] : "none";
                    var markerSvg = !string.IsNullOrEmpty(mShape) && mShape != "none"
                        ? RenderMarkerSvg(mShape, 8, 5, 3, color, color) : "";
                    var lineSvg = info.ScatterMarkersOnly
                        ? "" : $"<line x1=\"0\" y1=\"5\" x2=\"16\" y2=\"5\" stroke=\"{color}\" stroke-width=\"{lw:0.##}\"{dashAttr}/>";
                    swatch = $"<svg width=\"16\" height=\"10\" style=\"vertical-align:middle\">{lineSvg}{markerSvg}</svg>";
                }
                else
                {
                    swatch = $"<span style=\"display:inline-block;width:12px;height:12px;background:{color};border-radius:1px\"></span>";
                }
                sb.Append($"<span style=\"display:inline-flex;align-items:center;gap:4px\">{swatch}{HtmlEncode(info.Series[i].name)}</span>");
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

    /// <summary>SVG attribute fragment (with leading space) for the value-axis
    /// major-gridline dash, or "" when solid/absent. Driven by &lt;c:valAx&gt;
    /// &lt;c:majorGridlines&gt;&lt;c:spPr&gt;&lt;a:ln&gt;&lt;a:prstDash&gt;.</summary>
    private string ValGridDashAttr =>
        !string.IsNullOrEmpty(GridlineDash) && GridlineDash != "solid"
            ? $" stroke-dasharray=\"{RefLineDashArray(GridlineDash)}\""
            : "";

    // ==================== 3D Chart Helpers ====================

    /// <summary>Darken or lighten a hex color by a factor (0.0-2.0, 1.0=unchanged)</summary>
    // fill = marker interior (from <c:marker><c:spPr> solidFill, or series color),
    // stroke = marker outline (from <c:marker><c:spPr><a:ln>, or series color).
    // PowerPoint always draws a series-colored outline, so a white-filled marker
    // stays visible (hollow) on a white slide instead of vanishing.
    private static string RenderMarkerSvg(string shape, double cx, double cy, double r, string fill, string stroke)
    {
        // line-style glyphs (x/plus/dash/dot) have no interior — they are drawn
        // in the stroke color only; the fill arg is meaningless for them.
        const string sw = "1";
        return shape switch
        {
            "diamond" => $"<polygon points=\"{cx},{cy - r} {cx + r},{cy} {cx},{cy + r} {cx - r},{cy}\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\"/>",
            "square" => $"<rect x=\"{cx - r}\" y=\"{cy - r}\" width=\"{r * 2}\" height=\"{r * 2}\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\"/>",
            "triangle" => $"<polygon points=\"{cx},{cy - r} {cx + r},{cy + r} {cx - r},{cy + r}\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\"/>",
            "star" => BuildStarPath(cx, cy, r, fill, stroke),
            "x" => $"<g stroke=\"{stroke}\" stroke-width=\"1.5\"><line x1=\"{cx - r}\" y1=\"{cy - r}\" x2=\"{cx + r}\" y2=\"{cy + r}\"/><line x1=\"{cx + r}\" y1=\"{cy - r}\" x2=\"{cx - r}\" y2=\"{cy + r}\"/></g>",
            "plus" => $"<g stroke=\"{stroke}\" stroke-width=\"1.5\"><line x1=\"{cx}\" y1=\"{cy - r}\" x2=\"{cx}\" y2=\"{cy + r}\"/><line x1=\"{cx - r}\" y1=\"{cy}\" x2=\"{cx + r}\" y2=\"{cy}\"/></g>",
            "dash" => $"<line x1=\"{cx - r}\" y1=\"{cy}\" x2=\"{cx + r}\" y2=\"{cy}\" stroke=\"{stroke}\" stroke-width=\"2\"/>",
            "dot" => $"<circle cx=\"{cx}\" cy=\"{cy}\" r=\"1.5\" fill=\"{stroke}\"/>",
            "none" => "",
            _ => $"<circle cx=\"{cx}\" cy=\"{cy}\" r=\"{r}\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\"/>", // circle or auto
        };
    }

    private static string BuildStarPath(double cx, double cy, double r, string fill, string stroke)
    {
        var sb = new StringBuilder();
        sb.Append($"<polygon points=\"");
        for (int i = 0; i < 10; i++)
        {
            var angle = Math.PI / 2 + i * Math.PI / 5;
            var rad = i % 2 == 0 ? r : r * 0.4;
            sb.Append($"{cx + rad * Math.Cos(angle):0.#},{cy - rad * Math.Sin(angle):0.#} ");
        }
        sb.Append($"\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"1\"/>");
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
            if (ShowValGridlines)
            for (int t = 1; t <= tickCount; t++)
            {
                var gx = plotOx + (double)plotPw * t / tickCount;
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"{GridlineWidthPx:0.##}\"{ValGridDashAttr}/>");
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
                sb.AppendLine($"        <text x=\"{plotOx - 4}\" y=\"{oy + c * groupH + groupH / 2:0.#}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\"{CatTickWeightAttr} text-anchor=\"end\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= tickCount; t++)
            {
                var val = minVal + majorUnit * t;
                var label = FmtValAxis(val, valNumFmt);
                sb.AppendLine($"        <text x=\"{plotOx + (double)plotPw * t / tickCount:0.#}\" y=\"{oy + ph + 16}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\"{ValTickWeightAttr} text-anchor=\"middle\">{label}</text>");
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
            if (ShowValGridlines)
            for (int t = 1; t <= tickCount; t++)
            {
                var gy = oy + ph - (double)ph * t / tickCount;
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"{GridlineWidthPx:0.##}\"{ValGridDashAttr}/>");
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
                            sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{bx + barW / 2:0.#}\" y=\"{by + segH / 2:0.#}\" fill=\"white\" font-size=\"{DataLabelFontPx}\" text-anchor=\"middle\" dominant-baseline=\"middle\">{vlabel}</text>");
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
                            sb.AppendLine($"        <text class=\"chart-data-label\" x=\"{bx + barW / 2 + dx3d / 2:0.#}\" y=\"{by + dy3d - 3:0.#}\" fill=\"{DataLabelFill}\" font-size=\"{DataLabelFontPx}\"{DataLabelStyleAttr} text-anchor=\"middle\">{vlabel}</text>");
                        }
                    }
                }
            }
            // Category labels
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                sb.AppendLine($"        <text x=\"{ox + c * groupW + groupW / 2:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\"{CatTickWeightAttr} text-anchor=\"middle\">{HtmlEncode(label)}</text>");
            }
            // Value axis labels
            for (int t = 0; t <= tickCount; t++)
            {
                var val = minVal + majorUnit * t;
                var label = FmtValAxis(val, valNumFmt);
                var ty = oy + ph - ((val - minVal) / range) * ph;
                sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\"{ValTickWeightAttr} text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
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
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\"{CatTickWeightAttr} text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }

        // Y-axis value labels
        for (int t = 0; t <= 4; t++)
        {
            var val = maxVal * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\"{ValTickWeightAttr} text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
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
        if (ShowValGridlines)
        for (int t = 1; t <= 4; t++)
        {
            var gy = oy + plotH - (double)plotH * t / 4;
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + plotW}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"{GridlineWidthPx:0.##}\"{ValGridDashAttr}/>");
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
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + plotH + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\"{CatTickWeightAttr} text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        // Value axis
        for (int t = 0; t <= 4; t++)
        {
            var val = maxVal * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + plotH - (double)plotH * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\"{ValTickWeightAttr} text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
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
