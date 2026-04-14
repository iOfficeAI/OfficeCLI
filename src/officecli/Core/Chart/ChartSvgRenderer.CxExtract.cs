// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OfficeCli.Core;

/// <summary>
/// Extract ChartInfo from a cx:chart (Office 2016 extended chart) element and
/// emit SVG for the shape primitives that don't map onto the regular cChart
/// renderers (treemap nested rectangles, sunburst arcs, box-whisker boxes).
///
/// Histogram and funnel reuse the existing RenderBarChartSvg pipeline by
/// client-side binning (histogram) or treating the levels as categories
/// (funnel). Treemap / sunburst / boxWhisker have dedicated inline emitters.
///
/// This partial is on the same ChartSvgRenderer class so we have access to
/// the private helpers (HtmlEncode, colors, etc.).
/// </summary>
internal partial class ChartSvgRenderer
{
    // ==================== cx → ChartInfo extraction ====================

    /// <summary>
    /// Extract a <see cref="ChartInfo"/> from a cx:chart element. Produces
    /// the same shape the regular <c>ExtractChartInfo</c> does, so all of
    /// RenderChartSvgContent's downstream emitters work without branching on
    /// source format — except for the cx-specific types (treemap / sunburst /
    /// boxWhisker) which dispatch to new dedicated emitters in
    /// RenderChartSvgContent.
    /// </summary>
    public static ChartInfo ExtractCxChartInfo(CX.Chart chart)
    {
        var info = new ChartInfo();

        // ---- Title ----
        var chartTitle = chart.GetFirstChild<CX.ChartTitle>();
        if (chartTitle != null)
        {
            var titleText = chartTitle.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
            if (!string.IsNullOrEmpty(titleText)) info.Title = titleText;
            var titleRpr = chartTitle.Descendants<Drawing.RunProperties>().FirstOrDefault();
            if (titleRpr?.FontSize?.HasValue == true)
                info.TitleFontSize = $"{titleRpr.FontSize.Value / 100.0}pt";
            var titleColor = titleRpr?.GetFirstChild<Drawing.SolidFill>()
                ?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (!string.IsNullOrEmpty(titleColor)) info.TitleFontColor = $"#{titleColor}";
        }

        // ---- Series (plot area region) ----
        var plotArea = chart.GetFirstChild<CX.PlotArea>();
        var plotAreaRegion = plotArea?.GetFirstChild<CX.PlotAreaRegion>();
        var allSeries = plotAreaRegion?.Elements<CX.Series>().ToList() ?? new List<CX.Series>();
        var chartSpace = chart.Ancestors<CX.ChartSpace>().FirstOrDefault();
        var chartData = chartSpace?.GetFirstChild<CX.ChartData>();

        // Determine normalized chart type from the first series' LayoutId.
        // CX.SeriesLayout is a struct, not a C# enum, so we can't pattern-match
        // the typed value directly — compare via InnerText.
        var firstLayoutId = allSeries.FirstOrDefault()?.LayoutId?.InnerText ?? "";
        info.ChartType = firstLayoutId.ToLowerInvariant() switch
        {
            "funnel" => "funnel",
            "treemap" => "treemap",
            "sunburst" => "sunburst",
            "boxwhisker" => "boxwhisker",
            "clusteredcolumn" => "histogram",  // histogram is stored as clusteredColumn layoutId
            _ => "histogram"
        };

        // Read each series' data from the matching cx:data block (dataId.val → data.id).
        foreach (var series in allSeries)
        {
            var dataIdVal = series.GetFirstChild<CX.DataId>()?.Val?.Value ?? 0;
            var dataBlock = chartData?.Elements<CX.Data>().FirstOrDefault(d => (d.Id?.Value ?? 0) == dataIdVal);
            if (dataBlock == null) continue;

            var seriesName = series.GetFirstChild<CX.Text>()
                ?.GetFirstChild<CX.TextData>()
                ?.GetFirstChild<CX.VXsdstring>()?.Text ?? "Series";

            var values = dataBlock.Elements<CX.NumericDimension>()
                .SelectMany(nd => nd.Descendants<CX.NumericValue>())
                .Select(nv => double.TryParse(nv.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0.0)
                .ToArray();

            // Categories: strDim if present (funnel/treemap/sunburst), else values themselves (histogram)
            if (info.Categories.Length == 0)
            {
                var catStrDim = dataBlock.Elements<CX.StringDimension>()
                    .FirstOrDefault(sd => sd.Type?.Value == CX.StringDimensionType.Cat);
                if (catStrDim != null)
                {
                    info.Categories = catStrDim.Descendants<CX.ChartStringValue>()
                        .Select(cv => cv.Text ?? "")
                        .ToArray();
                }
            }

            info.Series.Add((seriesName, values));

            // Series fill color
            var spPrFill = series.GetFirstChild<CX.ShapeProperties>()
                ?.GetFirstChild<Drawing.SolidFill>()
                ?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (!string.IsNullOrEmpty(spPrFill)) info.Colors.Add($"#{spPrFill}");
        }

        // Fill in fallback colors for any series without an explicit spPr
        while (info.Colors.Count < info.Series.Count)
            info.Colors.Add(FallbackColors[info.Colors.Count % FallbackColors.Length]);

        // ---- Histogram-specific: bin the raw values into columns ----
        if (info.ChartType == "histogram" && info.Series.Count > 0)
        {
            var firstSeries = info.Series[0];
            var binning = allSeries.FirstOrDefault()?.Descendants<CX.Binning>().FirstOrDefault();
            var binCount = ReadBinCount(binning);
            var binSize = ReadBinSize(binning);
            var (binEdges, binCounts) = ComputeBins(firstSeries.values, binCount, binSize);
            // Replace values with bin counts, and categories with bin labels
            var labels = new string[binCounts.Length];
            for (int i = 0; i < binCounts.Length; i++)
            {
                var lo = FormatNumber(binEdges[i]);
                var hi = FormatNumber(binEdges[i + 1]);
                labels[i] = $"[{lo},{hi}]";
            }
            info.Categories = labels;
            info.Series[0] = (firstSeries.name, binCounts.Select(c => (double)c).ToArray());
            info.GapWidth = 0;  // histogram default — overridden below if cx:catScaling/@gapWidth is present
        }

        // ---- Axes: titles, scaling, styling ----
        //
        // Extracts the full per-axis vocabulary so it matches what the cx
        // builder emits (ChartExBuilder.BuildCategoryAxis / BuildValueAxis):
        //   - axismin/axismax/majorunit → cx:valScaling @min/@max/@majorUnit
        //   - gapWidth                  → cx:catScaling @gapWidth
        //   - gridlineColor             → cx:axis/cx:majorGridlines/cx:spPr/a:ln
        //   - axisline                  → cx:axis/cx:spPr/a:ln
        //   - axisfont (size+color)     → cx:axis/cx:txPr/.../a:defRPr
        //   - axis title font/bold      → cx:axis/cx:title/.../a:rPr
        //
        // Without these reads, any histogram that sets locked Y scale, custom
        // gridline/axis-line color, custom tick-label font, or custom axis
        // title bold/size renders in the HTML preview with Excel-default
        // values even though the XML is correct. Excel itself renders them
        // fine — this only affects officecli's in-process preview.
        if (plotArea != null)
        {
            var axes = plotArea.Elements<CX.Axis>().ToList();
            var catAxis = axes.FirstOrDefault();   // Id=0
            var valAxis = axes.ElementAtOrDefault(1);

            info.CatAxisTitle = ExtractAxisTitleText(catAxis);
            info.ValAxisTitle = ExtractAxisTitleText(valAxis);

            if (valAxis != null)
            {
                // Axis scaling (min/max/majorUnit) — string attributes on cx:valScaling.
                var valScaling = valAxis.GetFirstChild<CX.ValueAxisScaling>();
                if (valScaling != null)
                {
                    if (double.TryParse(valScaling.Min?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var mnV))
                        info.AxisMin = mnV;
                    if (double.TryParse(valScaling.Max?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var mxV))
                        info.AxisMax = mxV;
                    if (double.TryParse(valScaling.MajorUnit?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var muV))
                        info.MajorUnit = muV;
                }

                // Axis title font size / bold
                var valTitleEl = valAxis.Elements().FirstOrDefault(e => e.LocalName == "title");
                var valTitleRPr = valTitleEl?.Descendants<Drawing.RunProperties>().FirstOrDefault();
                if (valTitleRPr?.FontSize?.HasValue == true)
                    info.ValAxisTitleFontPx = (int)(valTitleRPr.FontSize.Value / 100.0);
                if (valTitleRPr?.Bold?.Value == true)
                    info.ValAxisTitleBold = true;

                // Tick label font — cx:axis/cx:txPr/.../a:defRPr (axisfont compound knob)
                var valTxPr = valAxis.Elements().FirstOrDefault(e => e.LocalName == "txPr");
                var valDefRPr = valTxPr?.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
                if (valDefRPr?.FontSize?.HasValue == true)
                    info.ValFontPx = (int)(valDefRPr.FontSize.Value / 100.0);
                info.ValFontColor = ExtractFontColor(valDefRPr);

                // Major gridline color
                var valGl = valAxis.Elements().FirstOrDefault(e => e.LocalName == "majorGridlines");
                var valGlSpPr = valGl?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                info.GridlineColor = ExtractLineColor(valGlSpPr);

                // Axis spine color
                var valSpPr = valAxis.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                info.AxisLineColor = ExtractLineColor(valSpPr);
            }

            if (catAxis != null)
            {
                // gapWidth — string attribute on cx:catScaling (overrides the
                // histogram default of 0 set during binning above).
                var catScaling = catAxis.GetFirstChild<CX.CategoryAxisScaling>();
                if (catScaling?.GapWidth?.Value is string gwStr
                    && int.TryParse(gwStr, out var gw))
                    info.GapWidth = gw;

                // Axis title font size / bold
                var catTitleEl = catAxis.Elements().FirstOrDefault(e => e.LocalName == "title");
                var catTitleRPr = catTitleEl?.Descendants<Drawing.RunProperties>().FirstOrDefault();
                if (catTitleRPr?.FontSize?.HasValue == true)
                    info.CatAxisTitleFontPx = (int)(catTitleRPr.FontSize.Value / 100.0);
                if (catTitleRPr?.Bold?.Value == true)
                    info.CatAxisTitleBold = true;

                // Tick label font
                var catTxPr = catAxis.Elements().FirstOrDefault(e => e.LocalName == "txPr");
                var catDefRPr = catTxPr?.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
                if (catDefRPr?.FontSize?.HasValue == true)
                    info.CatFontPx = (int)(catDefRPr.FontSize.Value / 100.0);
                info.CatFontColor = ExtractFontColor(catDefRPr);

                // Category-axis spine color (cataxis.line / axisline) — if
                // only axisline was set, both axes received identical outlines;
                // we still read cat separately so per-axis overrides work.
                // valSpPr is preferred but if valAxis has none we fall back
                // to catAxis for AxisLineColor.
                if (info.AxisLineColor == null)
                {
                    var catSpPr = catAxis.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                    info.AxisLineColor = ExtractLineColor(catSpPr);
                }
            }
        }

        // ---- Data labels (histogram) ----
        //
        // cx attaches dLbls to the series, not the chart type element. Read
        // cx:series/cx:dataLabels/cx:visibility[@value] to decide whether
        // the bar chart renderer should draw value labels above each bar.
        var firstSeriesEl = allSeries.FirstOrDefault();
        var dLabelsEl = firstSeriesEl?.GetFirstChild<CX.DataLabels>();
        if (dLabelsEl != null)
        {
            var vis = dLabelsEl.GetFirstChild<CX.DataLabelVisibilities>();
            if (vis?.Value?.Value == true)
            {
                info.ShowDataLabels = true;
                info.ShowDataLabelVal = true;
            }
        }

        // ---- Plot-area / chart-area background fills ----
        // Mirrors the regular cChart path in ExtractChartInfo: read the
        // spPr direct child of <cx:plotArea> and of <cx:chartSpace> and pull
        // the a:solidFill/a:srgbClr value. ExtractFillColor uses LocalName
        // matching so it works across c: and cx: namespaces unchanged.
        //
        // Downstream, PlotFillColor is painted as a <rect> inside the chart
        // SVG (RenderChartSvgContent) and ChartFillColor is applied as a
        // `background:` style on the chart container div (ExcelHandler
        // HtmlPreview). Without these lines, cx histograms with
        // `plotareafill` / `chartareafill` render on a blank white page
        // even though the XML is perfectly correct — the fills only
        // surface in Excel itself.
        var plotSpPr = plotArea?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        info.PlotFillColor = ExtractFillColor(plotSpPr);
        var chartSpPr = chartSpace?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        info.ChartFillColor = ExtractFillColor(chartSpPr);

        // ---- Legend ----
        // Presence-based (cx omits the element entirely to hide the legend,
        // unlike c:legend which uses <c:delete val="1"/>).
        var legend = chart.GetFirstChild<CX.Legend>();
        info.HasLegend = legend != null;
        if (legend != null)
        {
            // legendfont — cx:legend/cx:txPr/.../a:defRPr — compound
            // "size:color:fontname" knob from the builder.
            var legendTxPr = legend.Elements().FirstOrDefault(e => e.LocalName == "txPr");
            var legendDefRPr = legendTxPr?.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
            if (legendDefRPr?.FontSize?.HasValue == true)
                info.LegendFontSize = $"{legendDefRPr.FontSize.Value / 100.0:0.##}pt";
            info.LegendFontColor = ExtractFontColor(legendDefRPr);
        }

        return info;
    }

    private static string? ExtractAxisTitleText(CX.Axis? axis)
    {
        var title = axis?.GetFirstChild<CX.AxisTitle>();
        if (title == null) return null;
        return title.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
    }

    // ==================== Histogram binning (client-side) ====================

    // The cx binning XML uses raw OpenXmlUnknownElement children (val attribute
    // workaround — see ChartExBuilder.cs notes). Read val attribute directly.
    private static uint? ReadBinCount(CX.Binning? binning)
    {
        if (binning == null) return null;
        foreach (var child in binning.ChildElements)
        {
            if (child.LocalName != "binCount") continue;
            var val = child.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "val").Value;
            if (uint.TryParse(val, out var n)) return n;
        }
        return null;
    }

    private static double? ReadBinSize(CX.Binning? binning)
    {
        if (binning == null) return null;
        foreach (var child in binning.ChildElements)
        {
            if (child.LocalName != "binSize") continue;
            var val = child.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "val").Value;
            if (double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out var w))
                return w;
        }
        return null;
    }

    /// <summary>
    /// Compute histogram bins from raw values. Matches Excel's semantics:
    /// - If binCount is set, divide [min, max] into N equal-width bins.
    /// - If binSize is set, width = binSize, bins anchored at min.
    /// - Else auto-bin using sqrt(N) rule, clamped to [5, 20].
    /// Right-closed intervals (a, b] — the default for Excel's histogram.
    /// </summary>
    private static (double[] edges, int[] counts) ComputeBins(double[] values, uint? binCount, double? binSize)
    {
        if (values.Length == 0) return (new[] { 0.0, 1.0 }, new[] { 0 });
        var min = values.Min();
        var max = values.Max();
        if (Math.Abs(max - min) < 1e-9) { max = min + 1; }

        int n;
        double width;
        if (binSize is double sz && sz > 0)
        {
            width = sz;
            n = (int)Math.Max(1, Math.Ceiling((max - min) / width));
        }
        else
        {
            n = binCount is uint bc && bc > 0
                ? (int)bc
                : (int)Math.Clamp(Math.Ceiling(Math.Sqrt(values.Length)), 5, 20);
            width = (max - min) / n;
        }

        var edges = new double[n + 1];
        for (int i = 0; i <= n; i++) edges[i] = min + width * i;
        edges[n] = max; // clamp last edge to exact max to avoid FP drift

        var counts = new int[n];
        foreach (var v in values)
        {
            // Right-closed: find first bin where v <= edges[i+1]
            var idx = 0;
            for (int i = 0; i < n; i++)
            {
                if (v <= edges[i + 1]) { idx = i; break; }
                idx = n - 1;
            }
            counts[idx]++;
        }
        return (edges, counts);
    }

    private static string FormatNumber(double v)
    {
        // Short display — use "G" format for compact values, no trailing zeros.
        if (Math.Abs(v) >= 1000) return v.ToString("F0", CultureInfo.InvariantCulture);
        if (Math.Abs(v - Math.Round(v)) < 1e-9) return v.ToString("F0", CultureInfo.InvariantCulture);
        return v.ToString("0.##", CultureInfo.InvariantCulture);
    }

    // ==================== cx-specific SVG emitters ====================

    /// <summary>
    /// Render a funnel chart as centered horizontal bars. Excel funnels are
    /// drawn bottom-to-top with the widest level at the top, so we reverse
    /// the series order and render each level as a bar whose width is
    /// proportional to its value. Simple but visually conveys the shape.
    /// </summary>
    public void RenderCxFunnelSvg(StringBuilder sb, ChartInfo info,
        int marginLeft, int marginTop, int plotW, int plotH)
    {
        if (info.Series.Count == 0) return;
        var values = info.Series[0].values;
        var cats = info.Categories.Length == values.Length ? info.Categories : new string[values.Length];
        if (values.Length == 0) return;

        var maxVal = values.Max();
        if (maxVal <= 0) return;

        var rowH = (double)plotH / values.Length;
        var barH = rowH * 0.75;
        // Funnel: use a single series color (or first palette entry).
        // Cycling colors per level conflicts with the standard funnel look.
        var color = info.Colors.FirstOrDefault() ?? DefaultColors[0];
        var cx = marginLeft + plotW / 2;

        for (int i = 0; i < values.Length; i++)
        {
            var w = (values[i] / maxVal) * plotW;
            var y = marginTop + rowH * i + (rowH - barH) / 2;
            var x = cx - w / 2;
            sb.AppendLine($"    <rect x=\"{x:F1}\" y=\"{y:F1}\" width=\"{w:F1}\" height=\"{barH:F1}\" fill=\"{color}\" rx=\"2\"/>");
            // Label inside or to the right of bar
            var labelX = cx;
            var labelY = y + barH / 2;
            var label = (cats[i] ?? "") + $" ({FormatNumber(values[i])})";
            sb.AppendLine($"    <text x=\"{labelX}\" y=\"{labelY:F1}\" fill=\"#fff\" font-size=\"{CatFontPx}\" text-anchor=\"middle\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
        }
    }

    /// <summary>
    /// Render a treemap as a simple squarified layout. Treats all leaves as a
    /// flat list (ignores hierarchy — good enough for preview). Each rectangle's
    /// area is proportional to its value.
    ///
    /// Uses Bruls/Huijbregts/van Wijk (2000) squarify with row-wise fallback:
    /// pack items into strips along the shorter axis, finishing one strip
    /// before starting the next.
    /// </summary>
    public void RenderCxTreemapSvg(StringBuilder sb, ChartInfo info,
        int marginLeft, int marginTop, int plotW, int plotH)
    {
        if (info.Series.Count == 0) return;
        var values = info.Series[0].values;
        var cats = info.Categories.Length == values.Length ? info.Categories : new string[values.Length];
        if (values.Length == 0) return;
        var total = values.Sum();
        if (total <= 0) return;

        // Sort descending so big rectangles go first
        var order = Enumerable.Range(0, values.Length)
            .Where(i => values[i] > 0)
            .OrderByDescending(i => values[i]).ToArray();

        // Scale values so that sum equals rectangle area — then we can talk
        // directly in pixel areas for each cell.
        var scale = (double)plotW * plotH / total;
        var scaledVals = order.Select(i => values[i] * scale).ToArray();

        // Treemap / sunburst / funnel have ONE series but N cells, so cycle
        // through the palette per cell rather than painting every cell the
        // same series color. Use the theme palette if available.
        var palette = DefaultColors.Length > 0 ? DefaultColors : FallbackColors;

        var rect = new Rect { X = marginLeft, Y = marginTop, W = plotW, H = plotH };
        Squarify(scaledVals, 0, rect, (idx, r) =>
        {
            var origIdx = order[idx];
            var color = palette[origIdx % palette.Length];
            sb.AppendLine($"    <rect x=\"{r.X:F1}\" y=\"{r.Y:F1}\" width=\"{r.W:F1}\" height=\"{r.H:F1}\" fill=\"{color}\" stroke=\"#fff\" stroke-width=\"1.5\"/>");
            if (r.W > 40 && r.H > 18)
            {
                var label = cats[origIdx] ?? "";
                sb.AppendLine($"    <text x=\"{r.X + r.W / 2:F1}\" y=\"{r.Y + r.H / 2:F1}\" fill=\"#fff\" font-size=\"{CatFontPx}\" text-anchor=\"middle\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
            }
        });
    }

    private struct Rect { public double X, Y, W, H; }

    /// <summary>
    /// Classic squarify algorithm (Bruls et al. 2000), simplified: greedily
    /// group items into strips along the shorter side of the remaining rect,
    /// committing the strip when adding one more item would worsen the worst
    /// aspect ratio of the current group. Each committed strip consumes the
    /// full shorter side; remaining items fill the leftover rectangle.
    /// </summary>
    private static void Squarify(double[] areas, int start, Rect rect, Action<int, Rect> emit)
    {
        if (start >= areas.Length || rect.W <= 0.5 || rect.H <= 0.5) return;

        // Convention: the "strip" is placed along the SHORT side. If the
        // rectangle is WIDE (W > H), the strip is a vertical column at the
        // left edge (full H tall, stripW wide). If the rectangle is TALL
        // (H > W), the strip is a horizontal row at the top edge (full W
        // wide, stripH tall). Items stack ALONG the short side (vertically
        // in a wide rect, horizontally in a tall rect).
        var shortSide = Math.Min(rect.W, rect.H);

        // Greedily extend the current row as long as aspect ratio improves
        // (or stays equal). Stop and commit when the next item would make
        // the worst aspect ratio worse.
        int end = start + 1;
        double bestWorst = RowWorstRatio(areas, start, end, shortSide);

        while (end < areas.Length)
        {
            var tryEnd = end + 1;
            var tryWorst = RowWorstRatio(areas, start, tryEnd, shortSide);
            if (tryWorst <= bestWorst)
            {
                end = tryEnd;
                bestWorst = tryWorst;
            }
            else break;
        }

        // Emit the committed row
        var stripAdvance = LayoutRow(areas, start, end, rect, emit);

        // Recurse on the leftover rectangle (the part outside the strip).
        Rect remaining = rect.W >= rect.H
            // Wide rect → vertical strip at left → recurse on right slab
            ? new Rect { X = rect.X + stripAdvance, Y = rect.Y, W = rect.W - stripAdvance, H = rect.H }
            // Tall rect → horizontal strip at top → recurse on bottom slab
            : new Rect { X = rect.X, Y = rect.Y + stripAdvance, W = rect.W, H = rect.H - stripAdvance };

        Squarify(areas, end, remaining, emit);
    }

    /// <summary>
    /// Worst aspect ratio for the proposed row (items [start, end)) packed
    /// along a strip of length <paramref name="shortSide"/>. Each item then
    /// has one dimension = stripThickness = rowSum/shortSide and the other
    /// = a_i / stripThickness. Per Bruls et al.:
    ///     worst = max(max_i(w² · a_max / s²), max_i(s² / (w² · a_min)))
    /// </summary>
    private static double RowWorstRatio(double[] areas, int start, int end, double shortSide)
    {
        if (end <= start) return double.MaxValue;
        double s = 0;
        double maxArea = 0, minArea = double.MaxValue;
        for (int i = start; i < end; i++)
        {
            s += areas[i];
            if (areas[i] > maxArea) maxArea = areas[i];
            if (areas[i] < minArea) minArea = areas[i];
        }
        if (s <= 0 || shortSide <= 0) return double.MaxValue;
        var sqSide = shortSide * shortSide;
        var a = (sqSide * maxArea) / (s * s);
        var b = (s * s) / (sqSide * Math.Max(minArea, 1e-9));
        return Math.Max(a, b);
    }

    /// <summary>
    /// Lay out a committed row inside <paramref name="rect"/> and call
    /// <paramref name="emit"/> for each item. Returns how far the strip
    /// advanced along the LONG side of the rectangle — the caller uses
    /// this to compute the leftover rectangle.
    /// </summary>
    private static double LayoutRow(double[] areas, int start, int end, Rect rect, Action<int, Rect> emit)
    {
        double rowSum = 0;
        for (int i = start; i < end; i++) rowSum += areas[i];
        if (rowSum <= 0) return 0;

        var wideRect = rect.W >= rect.H;
        var shortSide = Math.Min(rect.W, rect.H);
        var stripThickness = rowSum / shortSide;  // strip depth along long side

        // Items inside the strip have one fixed side = stripThickness and
        // the other side = a_i / stripThickness. They stack along the short
        // side of the original rect.
        var cursor = 0.0;
        for (int i = start; i < end; i++)
        {
            var itemLenAlongShort = areas[i] / stripThickness;
            Rect r;
            if (wideRect)
            {
                // Strip is a vertical column at rect.X, full height stacked.
                r = new Rect
                {
                    X = rect.X,
                    Y = rect.Y + cursor,
                    W = stripThickness,
                    H = itemLenAlongShort,
                };
            }
            else
            {
                // Strip is a horizontal row at rect.Y, full width packed.
                r = new Rect
                {
                    X = rect.X + cursor,
                    Y = rect.Y,
                    W = itemLenAlongShort,
                    H = stripThickness,
                };
            }
            emit(i, r);
            cursor += itemLenAlongShort;
        }
        return stripThickness;
    }

    /// <summary>
    /// Render a sunburst as concentric arcs. Without full hierarchy info we
    /// just draw a single ring with one slice per value (like a pie chart
    /// with a large hole). Good enough for previews.
    /// </summary>
    public void RenderCxSunburstSvg(StringBuilder sb, ChartInfo info,
        int marginLeft, int marginTop, int plotW, int plotH)
    {
        if (info.Series.Count == 0) return;
        var values = info.Series[0].values;
        var cats = info.Categories.Length == values.Length ? info.Categories : new string[values.Length];
        var total = values.Sum();
        if (total <= 0) return;

        var cx = marginLeft + plotW / 2.0;
        var cy = marginTop + plotH / 2.0;
        var rOuter = Math.Min(plotW, plotH) / 2.0 - 10;
        var rInner = rOuter * 0.35;

        var palette = DefaultColors.Length > 0 ? DefaultColors : FallbackColors;
        var startAngle = -Math.PI / 2; // start at 12 o'clock
        for (int i = 0; i < values.Length; i++)
        {
            var sweep = (values[i] / total) * 2 * Math.PI;
            if (sweep <= 0) continue;
            var endAngle = startAngle + sweep;
            var largeArc = sweep > Math.PI ? 1 : 0;

            var x1 = cx + rOuter * Math.Cos(startAngle);
            var y1 = cy + rOuter * Math.Sin(startAngle);
            var x2 = cx + rOuter * Math.Cos(endAngle);
            var y2 = cy + rOuter * Math.Sin(endAngle);
            var ix1 = cx + rInner * Math.Cos(endAngle);
            var iy1 = cy + rInner * Math.Sin(endAngle);
            var ix2 = cx + rInner * Math.Cos(startAngle);
            var iy2 = cy + rInner * Math.Sin(startAngle);

            var d = $"M {x1:F1},{y1:F1} A {rOuter:F1},{rOuter:F1} 0 {largeArc} 1 {x2:F1},{y2:F1} "
                  + $"L {ix1:F1},{iy1:F1} A {rInner:F1},{rInner:F1} 0 {largeArc} 0 {ix2:F1},{iy2:F1} Z";
            var color = palette[i % palette.Length];
            sb.AppendLine($"    <path d=\"{d}\" fill=\"{color}\" stroke=\"#fff\" stroke-width=\"1\"/>");

            // Label in the middle of the arc
            var midAngle = startAngle + sweep / 2;
            var labelR = (rOuter + rInner) / 2;
            var lx = cx + labelR * Math.Cos(midAngle);
            var ly = cy + labelR * Math.Sin(midAngle);
            var label = cats[i] ?? "";
            if (sweep > 0.25 && !string.IsNullOrEmpty(label))
                sb.AppendLine($"    <text x=\"{lx:F1}\" y=\"{ly:F1}\" fill=\"#fff\" font-size=\"{CatFontPx}\" text-anchor=\"middle\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");

            startAngle = endAngle;
        }
    }

    /// <summary>
    /// Render a box-whisker chart. For each series: box (Q1–Q3), median line,
    /// whiskers extending to the last non-outlier value within 1.5×IQR of the
    /// fence, outlier data points drawn as open circles, and a mean marker (×).
    /// </summary>
    public void RenderCxBoxWhiskerSvg(StringBuilder sb, ChartInfo info,
        int marginLeft, int marginTop, int plotW, int plotH)
    {
        if (info.Series.Count == 0) return;

        // Compute stats per series
        var stats = info.Series.Select(s => ComputeBoxStats(s.values)).ToList();
        if (stats.All(s => s == null)) return;

        // Global scale includes outliers
        var globalMin = stats.Where(s => s != null).Min(s => s!.Value.allMin);
        var globalMax = stats.Where(s => s != null).Max(s => s!.Value.allMax);
        if (Math.Abs(globalMax - globalMin) < 1e-9) globalMax = globalMin + 1;
        // Add 5% padding so top/bottom outliers aren't clipped at the edge
        var pad = (globalMax - globalMin) * 0.05;
        globalMin -= pad;
        globalMax += pad;

        var bw = (double)plotW / info.Series.Count;
        var boxW = bw * 0.5;

        double yCoord(double v) => marginTop + plotH - ((v - globalMin) / (globalMax - globalMin)) * plotH;

        // Y axis: a few tick labels for context
        for (int t = 0; t <= 4; t++)
        {
            var v = globalMin + pad + (globalMax - globalMin - 2 * pad) * t / 4;
            var y = yCoord(v);
            sb.AppendLine($"    <line x1=\"{marginLeft}\" y1=\"{y:F1}\" x2=\"{marginLeft + plotW}\" y2=\"{y:F1}\" stroke=\"{GridColor}\" stroke-dasharray=\"2,2\"/>");
            sb.AppendLine($"    <text x=\"{marginLeft - 3}\" y=\"{y:F1}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{FormatNumber(v)}</text>");
        }

        for (int si = 0; si < info.Series.Count; si++)
        {
            if (stats[si] is not { } s) continue;
            var color = info.Colors[si % info.Colors.Count];
            var cxCenter = marginLeft + bw * (si + 0.5);
            var boxX = cxCenter - boxW / 2;

            var yWLow  = yCoord(s.whiskerLow);
            var yWHigh = yCoord(s.whiskerHigh);
            var yQ1    = yCoord(s.q1);
            var yQ3    = yCoord(s.q3);
            var yMed   = yCoord(s.median);
            var yMean  = yCoord(s.mean);

            // Whisker vertical line: Q1→whiskerLow and Q3→whiskerHigh
            sb.AppendLine($"    <line x1=\"{cxCenter:F1}\" y1=\"{yWLow:F1}\" x2=\"{cxCenter:F1}\" y2=\"{yQ1:F1}\" stroke=\"{color}\" stroke-width=\"1.5\"/>");
            sb.AppendLine($"    <line x1=\"{cxCenter:F1}\" y1=\"{yQ3:F1}\" x2=\"{cxCenter:F1}\" y2=\"{yWHigh:F1}\" stroke=\"{color}\" stroke-width=\"1.5\"/>");
            // Whisker caps (horizontal ticks at fence endpoints)
            var capHalf = boxW * 0.3;
            sb.AppendLine($"    <line x1=\"{cxCenter - capHalf:F1}\" y1=\"{yWLow:F1}\" x2=\"{cxCenter + capHalf:F1}\" y2=\"{yWLow:F1}\" stroke=\"{color}\" stroke-width=\"1.5\"/>");
            sb.AppendLine($"    <line x1=\"{cxCenter - capHalf:F1}\" y1=\"{yWHigh:F1}\" x2=\"{cxCenter + capHalf:F1}\" y2=\"{yWHigh:F1}\" stroke=\"{color}\" stroke-width=\"1.5\"/>");
            // Box Q1..Q3
            sb.AppendLine($"    <rect x=\"{boxX:F1}\" y=\"{yWHigh:F1}\" width=\"{boxW:F1}\" height=\"{yWLow - yWHigh:F1}\" fill=\"{color}\" fill-opacity=\"0.25\" stroke=\"{color}\" stroke-width=\"1.5\"/>");
            // Median line
            sb.AppendLine($"    <line x1=\"{boxX:F1}\" y1=\"{yMed:F1}\" x2=\"{boxX + boxW:F1}\" y2=\"{yMed:F1}\" stroke=\"{color}\" stroke-width=\"2.5\"/>");
            // Mean marker: × symbol
            var mx = 4.0;
            sb.AppendLine($"    <line x1=\"{cxCenter - mx:F1}\" y1=\"{yMean - mx:F1}\" x2=\"{cxCenter + mx:F1}\" y2=\"{yMean + mx:F1}\" stroke=\"{color}\" stroke-width=\"1.5\"/>");
            sb.AppendLine($"    <line x1=\"{cxCenter + mx:F1}\" y1=\"{yMean - mx:F1}\" x2=\"{cxCenter - mx:F1}\" y2=\"{yMean + mx:F1}\" stroke=\"{color}\" stroke-width=\"1.5\"/>");

            // Outlier circles
            const double r = 3.5;
            foreach (var ov in s.outliers)
            {
                var yo = yCoord(ov);
                sb.AppendLine($"    <circle cx=\"{cxCenter:F1}\" cy=\"{yo:F1}\" r=\"{r}\" fill=\"none\" stroke=\"{color}\" stroke-width=\"1.2\"/>");
            }

            // Series label
            sb.AppendLine($"    <text x=\"{cxCenter:F1}\" y=\"{marginTop + plotH + 14}\" fill=\"{AxisColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(info.Series[si].name)}</text>");
        }
    }

    private record struct BoxStats(
        double whiskerLow, double q1, double median, double q3, double whiskerHigh,
        double mean, double allMin, double allMax, double[] outliers);

    private static BoxStats? ComputeBoxStats(double[] values)
    {
        if (values.Length == 0) return null;
        var sorted = values.OrderBy(v => v).ToArray();
        double Percentile(double p)
        {
            if (sorted.Length == 1) return sorted[0];
            var idx = p * (sorted.Length - 1);
            var lo = (int)Math.Floor(idx);
            var hi = (int)Math.Ceiling(idx);
            var frac = idx - lo;
            return sorted[lo] * (1 - frac) + sorted[hi] * frac;
        }
        var q1 = Percentile(0.25);
        var q3 = Percentile(0.75);
        var iqr = q3 - q1;
        var fenceLow  = q1 - 1.5 * iqr;
        var fenceHigh = q3 + 1.5 * iqr;

        // Whiskers extend to the last data point within the fence
        var whiskerLow  = sorted.Where(v => v >= fenceLow).DefaultIfEmpty(q1).Min();
        var whiskerHigh = sorted.Where(v => v <= fenceHigh).DefaultIfEmpty(q3).Max();
        var outliers    = sorted.Where(v => v < fenceLow || v > fenceHigh).ToArray();
        var mean        = sorted.Average();

        return new BoxStats(
            whiskerLow, q1, Percentile(0.5), q3, whiskerHigh,
            mean, sorted[0], sorted[^1], outliers);
    }

    /// <summary>
    /// Dispatcher entry for cx chart types that aren't reducible to the
    /// regular bar/column pipeline. Histogram → RenderBarChartSvg (handled
    /// by the main dispatcher after ExtractCxChartInfo pre-bins the data).
    /// </summary>
    public bool TryRenderCxSpecificType(StringBuilder sb, ChartInfo info,
        int marginLeft, int marginTop, int plotW, int plotH)
    {
        switch (info.ChartType)
        {
            case "funnel":
                RenderCxFunnelSvg(sb, info, marginLeft, marginTop, plotW, plotH);
                return true;
            case "treemap":
                RenderCxTreemapSvg(sb, info, marginLeft, marginTop, plotW, plotH);
                return true;
            case "sunburst":
                RenderCxSunburstSvg(sb, info, marginLeft, marginTop, plotW, plotH);
                return true;
            case "boxwhisker":
                RenderCxBoxWhiskerSvg(sb, info, marginLeft, marginTop, plotW, plotH);
                return true;
        }
        return false;
    }
}
