// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

internal static partial class ChartHelper
{
    // ==================== Build ChartSpace ====================

    internal static C.ChartSpace BuildChartSpace(
        string chartType,
        string? title,
        string[]? categories,
        List<(string name, double[] values)> seriesData,
        Dictionary<string, string> properties)
    {
        var (kind, is3D, stacked, percentStacked) = ParseChartType(chartType);

        var chartSpace = new C.ChartSpace();
        // Always write an explicit <c:date1904> (first CT_ChartSpace child).
        // When absent, real PowerPoint may fall back to 1904-epoch date
        // interpretation — date-axis serials render four years off.
        chartSpace.AppendChild(new C.Date1904
        {
            Val = properties.TryGetValue("date1904", out var d1904)
                  && OfficeCli.Core.ParseHelpers.IsTruthy(d1904)
        });
        var chart = new C.Chart();

        // BUG-001: `title=none` (case-insensitive) or empty string means "no
        // title" — do not render the literal text "none". Mirrors the Setter's
        // `title` handling in ChartHelper.Setter.cs.
        if (!string.IsNullOrEmpty(title) && !title.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            // R53 tester-2: forward an explicit title.lang (defaults to en-US
            // inside BuildChartTitle when absent) so dump→replay preserves the
            // source locale on the chart-title run.
            properties.TryGetValue("title.lang", out var titleLangBuild);
            OfficeCli.Core.ParseHelpers.ValidateXmlText(title, "title");
            chart.AppendChild(BuildChartTitle(title, titleLangBuild));
        }
        else if (properties.TryGetValue("autoTitle", out var autoTitleVal)
                 && OfficeCli.Core.ParseHelpers.IsTruthy(autoTitleVal))
        {
            // Empty <c:title/> + autoTitleDeleted=0 → real PowerPoint renders
            // its localized "Chart Title" placeholder (multi-series source
            // charts authored with an automatic title). The captured
            // title.pPr (font size/color) rides in c:title/c:txPr so the
            // placeholder keeps the authored styling.
            var autoTitleEl = new C.Title();
            if (properties.TryGetValue("title.pPr", out var autoTitlePPr)
                && !string.IsNullOrWhiteSpace(autoTitlePPr))
            {
                try
                {
                    var autoTitlePara = new Drawing.Paragraph();
                    autoTitlePara.AppendChild(new Drawing.ParagraphProperties(autoTitlePPr));
                    var autoTitleTxPr = new C.TextProperties(
                        new Drawing.BodyProperties(), new Drawing.ListStyle());
                    autoTitleTxPr.AppendChild(autoTitlePara);
                    autoTitleEl.AppendChild(autoTitleTxPr);
                }
                catch { /* malformed captured pPr — placeholder stays unstyled */ }
            }
            chart.AppendChild(autoTitleEl);
            chart.AppendChild(new C.AutoTitleDeleted { Val = false });
        }

        var originalCategories = categories;

        // R46 Major-3: pie / doughnut / pieOfPie / barOfPie all render exactly
        // one series. When user passes `data=Name1:V1;Name2:V2;...` the splitter
        // produces N single-value "series", and only seriesData[0] is consumed
        // downstream — points 2..N are silently dropped. Coalesce here so the
        // ; form means "N data points in one series" (categories = names).
        // Mirrors the waterfall flattener at line ~214 in this file.
        if ((kind is "pie" or "doughnut" or "pieofpie" or "barofpie")
            && seriesData.Count > 1
            && seriesData.All(s => s.values.Length == 1))
        {
            // Coalesce N single-value series into one N-point series. Mirrors
            // the waterfall flattener below: the data points are always merged;
            // user-supplied categories (e.g. `categories="A,B,C,D,E"`) label the
            // points and are kept as-is. Only when no categories were given do we
            // derive them from the per-series names. Previously this whole block
            // was gated on `originalCategories == null`, so passing categories
            // alongside `series1=..series5=` left 5 single-value series and the
            // doughnut/pie renderer consumed only seriesData[0] — a single
            // full-circle slice that draws as a degenerate (start==end) arc.
            if (originalCategories == null)
            {
                categories = seriesData.Select(s => s.name).ToArray();
                originalCategories = categories;
            }
            var coalesced = seriesData.SelectMany(s => s.values).ToArray();
            seriesData = new List<(string name, double[] values)>
            {
                ("Series 1", coalesced)
            };
        }

        if (categories == null && seriesData.Count > 0)
        {
            var maxLen = seriesData.Max(s => s.values.Length);
            categories = Enumerable.Range(1, maxLen).Select(i => i.ToString()).ToArray();
        }

        var plotArea = new C.PlotArea(new C.Layout());
        uint catAxisId = 1;
        uint valAxisId = 2;
        uint serAxisId = 3;

        OpenXmlCompositeElement? chartElement;
        bool needsAxes = true;
        // 3D charts (bar3D/column3D/line3D/area3D) require a third axId on the
        // chart element (CT_Bar3DChart / CT_Line3DChart / CT_Area3DChart) bound
        // to a c:serAx in the plotArea. Without it, OpenXmlValidator rejects the
        // chart and PowerPoint silently repairs the file with a degenerate
        // render (single ribbon / no series progression).
        bool needsSerAxis = false;

        var colors = ParseSeriesColors(properties);
        // BUG-DUMP-R36-3: series whose dump carried per-point dPt styling but no
        // series-level fill — suppress the default accent1 spPr for them.
        var noFillSeries = SeriesWithNoCapturedFill(properties);
        // CONSISTENCY(chart-series-color): for single-series chart kinds
        // (pie/doughnut/stock), the merged `colors[]` is consumed as per-point
        // dPt overrides. A user typing `series1.color=#FF0000` expects the
        // whole series tinted (matches `set chart` ApplySeriesColor), which
        // would otherwise be silently dropped. Capture the dotted-only spec so
        // those builders can call ApplySeriesColor in addition to per-point.
        var dottedSeriesColors = ParseDottedSeriesColorsOnly(properties);
        // Positional `colors=` list (per-data-point for single-series kinds),
        // SEPARATE from `series{N}.color` (whole-series tint). Doughnut /
        // pie / pie3D / pieofpie / barofpie use this so a user-supplied
        // `series1.color=#C00000` does NOT also emit a spurious dPt#0 with
        // the same fill (Get→dump→replay then surfaced both series1.color
        // and point1.color, breaking byte-equality round-trip).
        var pointColors = ParsePositionalColorsOnly(properties);

        switch (kind)
        {
            case "bar" when is3D:
            case "column" when is3D:
            {
                var dir3dAuto = kind == "bar" ? C.BarDirectionValues.Bar : C.BarDirectionValues.Column;
                var bar3dAuto = new C.Bar3DChart(
                    new C.BarDirection { Val = dir3dAuto },
                    new C.BarGrouping { Val = stacked ? C.BarGroupingValues.Stacked
                        : percentStacked ? C.BarGroupingValues.PercentStacked
                        : C.BarGroupingValues.Clustered },
                    new C.VaryColors { Val = false }
                );
                for (int si = 0; si < seriesData.Count; si++)
                {
                    var s = BuildBarSeries((uint)si, seriesData[si].name, categories, seriesData[si].values,
                        colors != null && si < colors.Length ? colors[si] : null);
                    bar3dAuto.AppendChild(s);
                }
                bar3dAuto.AppendChild(new C.GapWidth { Val = 150 });
                bar3dAuto.AppendChild(new C.AxisId { Val = catAxisId });
                bar3dAuto.AppendChild(new C.AxisId { Val = valAxisId });
                bar3dAuto.AppendChild(new C.AxisId { Val = serAxisId });
                chartElement = bar3dAuto;
                needsSerAxis = true;
                break;
            }
            case "bar":
                chartElement = BuildBarChart(C.BarDirectionValues.Bar, stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors, noFillSeries);
                break;
            case "column":
                chartElement = BuildBarChart(C.BarDirectionValues.Column, stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors, noFillSeries);
                break;
            case "line" when is3D:
            {
                var grouping3d = percentStacked ? C.GroupingValues.PercentStacked
                    : stacked ? C.GroupingValues.Stacked
                    : C.GroupingValues.Standard;
                var line3d = new C.Line3DChart(
                    new C.Grouping { Val = grouping3d },
                    new C.VaryColors { Val = false }
                );
                for (int i = 0; i < seriesData.Count; i++)
                {
                    var color = ResolveSeriesColor(colors, i, noFillSeries);
                    line3d.AppendChild(BuildLineSeries((uint)i, seriesData[i].name,
                        categories, seriesData[i].values, color));
                }
                line3d.AppendChild(new C.AxisId { Val = catAxisId });
                line3d.AppendChild(new C.AxisId { Val = valAxisId });
                line3d.AppendChild(new C.AxisId { Val = serAxisId });
                chartElement = line3d;
                needsSerAxis = true;
                break;
            }
            case "line":
                chartElement = BuildLineChart(stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors, noFillSeries);
                break;
            case "area" when is3D:
            {
                var grouping3d = percentStacked ? C.GroupingValues.PercentStacked
                    : stacked ? C.GroupingValues.Stacked
                    : C.GroupingValues.Standard;
                var area3d = new C.Area3DChart(
                    new C.Grouping { Val = grouping3d },
                    new C.VaryColors { Val = false }
                );
                for (int i = 0; i < seriesData.Count; i++)
                {
                    var color = ResolveSeriesColor(colors, i, noFillSeries);
                    area3d.AppendChild(BuildAreaSeries((uint)i, seriesData[i].name,
                        categories, seriesData[i].values, color));
                }
                area3d.AppendChild(new C.AxisId { Val = catAxisId });
                area3d.AppendChild(new C.AxisId { Val = valAxisId });
                area3d.AppendChild(new C.AxisId { Val = serAxisId });
                chartElement = area3d;
                needsSerAxis = true;
                break;
            }
            case "area":
                chartElement = BuildAreaChart(stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors, noFillSeries);
                break;
            case "pie" when is3D:
            {
                // Pie charts vary color by data point, not by series. Writing a
                // series-level solidFill (via BuildPieSeries(..., color)) overrides
                // varyColors and renders every slice in one color. Match the 2D
                // pie path: build the series without a series color, then attach
                // per-point colors through ApplyDataPointColors.
                var pie3d = new C.Pie3DChart(
                    new C.VaryColors { Val = true }
                );
                if (seriesData.Count > 0)
                {
                    var series = BuildPieSeries(0, seriesData[0].name,
                        categories, seriesData[0].values);
                    ApplyDataPointColors(series, seriesData[0].values.Length, pointColors);
                    ApplyDottedSeriesColors(new[] { (OpenXmlCompositeElement)series }, dottedSeriesColors);
                    pie3d.AppendChild(series);
                }
                chartElement = pie3d;
                needsAxes = false;
                break;
            }
            case "pie":
                chartElement = BuildPieChart(categories, seriesData, pointColors);
                ApplyDottedSeriesColors(((C.PieChart)chartElement).Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>().ToArray(), dottedSeriesColors);
                needsAxes = false;
                break;
            case "pieofpie":
            case "barofpie":
                chartElement = BuildOfPieChart(
                    kind == "barofpie" ? C.OfPieValues.Bar : C.OfPieValues.Pie,
                    categories, seriesData, pointColors);
                ApplyDottedSeriesColors(((C.OfPieChart)chartElement).Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>().ToArray(), dottedSeriesColors);
                needsAxes = false;
                break;
            case "doughnut":
                chartElement = BuildDoughnutChart(categories, seriesData, pointColors);
                ApplyDottedSeriesColors(((C.DoughnutChart)chartElement).Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>().ToArray(), dottedSeriesColors);
                needsAxes = false;
                break;
            case "scatter":
                var scatterStyle = properties.GetValueOrDefault("scatterStyle", "lineMarker");
                chartElement = BuildScatterChart(categories, seriesData, catAxisId, valAxisId, scatterStyle, colors);
                break;
            case "bubble":
                chartElement = BuildBubbleChart(categories, seriesData, catAxisId, valAxisId, colors, properties);
                break;
            case "radar":
            {
                var radarStyle = properties.GetValueOrDefault("radarStyle", "marker");
                chartElement = BuildRadarChart(radarStyle, categories, seriesData, catAxisId, valAxisId, colors);
                break;
            }
            // Note: column3d/bar3d are handled by "column when is3D" / "bar when is3D" above
            case "stock":
                chartElement = BuildStockChart(categories, seriesData, catAxisId, valAxisId);
                // Stock series default to NoFill outline (hi-lo lines carry the
                // visual). series{N}.color must still be honored — apply after
                // build so user-chosen tints reach the LineChartSeries spPr.
                ApplyDottedSeriesColors(((C.StockChart)chartElement).Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>().ToArray(), dottedSeriesColors);
                needsAxes = true;
                break;
            case "waterfall":
            {
                // Waterfall chart via stacked bar simulation
                double[] wfValues;
                string[]? wfCategories = categories;

                if (seriesData.Count > 1 && seriesData.All(s => s.values.Length == 1))
                {
                    // User passed per-category name:value format (e.g. "Start:1000,Revenue:500,Expense:-200,Net:1300")
                    // Flatten: use series names as categories, combine all single values into one array
                    if (originalCategories == null)
                        wfCategories = seriesData.Select(s => s.name).ToArray();
                    wfValues = seriesData.Select(s => s.values[0]).ToArray();
                }
                else
                {
                    wfValues = seriesData.Count > 0 ? seriesData[0].values : Array.Empty<double>();
                }

                var incColor = properties.GetValueOrDefault("increaseColor");
                var decColor = properties.GetValueOrDefault("decreaseColor");
                var totColor = properties.GetValueOrDefault("totalColor");
                var wfChartSpace = BuildWaterfallChart(title, wfCategories, wfValues,
                    incColor, decColor, totColor, properties);
                // BuildWaterfallChart builds its own chart with a default
                // legend and never consumed the `legend` prop, so legend= was
                // reported unsupported on a waterfall Add (every other type
                // accepts it). Drop the default legend and re-apply from the
                // prop (position, or none) — remove-first avoids a duplicate
                // <c:legend> that would make Excel refuse the file.
                var wfChart = wfChartSpace.GetFirstChild<C.Chart>();
                if (wfChart != null)
                {
                    wfChart.RemoveAllChildren<C.Legend>();
                    ApplyLegendFromProps(wfChart, properties);
                }
                return wfChartSpace;
            }
            case "combo":
            {
                // Reader emits per-series type list as `comboTypes=column,column,line`
                // (or area,area,column,column,line — any mix). Honor it so dump→replay
                // of line+area / bar+line / area+column+line combos rebuilds the
                // correct CT_*Chart elements with each ser bound to its original type.
                // Falls back to the legacy column-vs-line split at `combosplit` when
                // comboTypes is absent.
                string[]? comboTypes = null;
                if (properties.TryGetValue("comboTypes", out var ctList))
                    comboTypes = ctList.Split(',').Select(t => t.Trim().ToLowerInvariant()).ToArray();
                // Probe combosplit unconditionally: it is a legitimately
                // accepted prop (dump emits it alongside comboTypes), and the
                // TrackingPropertyDictionary reports any never-read key as
                // unsupported_property — reading it only inside the fallback
                // branch produced a false warning on every combo-chart replay.
                var hasComboSplit = properties.TryGetValue("combosplit", out var splitStr);
                if (comboTypes == null || comboTypes.Length == 0)
                {
                    int splitAt = 1;
                    if (hasComboSplit)
                        splitAt = ParseHelpers.SafeParseInt(splitStr!, "combosplit");
                    splitAt = Math.Min(splitAt, seriesData.Count);
                    comboTypes = new string[seriesData.Count];
                    for (int i = 0; i < seriesData.Count; i++)
                        comboTypes[i] = i < splitAt ? "column" : "line";
                }

                // Group adjacent series of identical type into one CT_*Chart container
                // (matches OOXML structure: each chart element holds a contiguous
                // run of series of its own type; combo charts may interleave by
                // chart-element ordering but each container is homogeneous).
                int gi = 0;
                while (gi < seriesData.Count)
                {
                    var t = gi < comboTypes.Length ? comboTypes[gi] : "line";
                    int gj = gi + 1;
                    while (gj < seriesData.Count
                           && gj < comboTypes.Length
                           && comboTypes[gj] == t)
                        gj++;
                    BuildComboGroup(t, plotArea, seriesData, categories,
                        startIdx: gi, endIdxExclusive: gj,
                        catAxisId, valAxisId, colors, noFillSeries);
                    gi = gj;
                }
                chartElement = null;
                break;
            }
            default:
                throw new ArgumentException(
                    $"Unknown chart type: '{kind}'. Supported: column, bar, line, pie, doughnut, area, scatter, bubble, radar, stock, combo, waterfall. " +
                    "Add 'stacked' or 'percentstacked' suffix for variants (e.g. columnstacked).");
        }

        if (chartElement != null)
            plotArea.AppendChild(chartElement);

        if (needsAxes)
        {
            if (kind == "scatter")
            {
                plotArea.AppendChild(BuildValueAxis(catAxisId, valAxisId, C.AxisPositionValues.Bottom));
                plotArea.AppendChild(BuildValueAxis(valAxisId, catAxisId, C.AxisPositionValues.Left));
            }
            else
            {
                plotArea.AppendChild(BuildCategoryAxis(catAxisId, valAxisId));
                plotArea.AppendChild(BuildValueAxis(valAxisId, catAxisId, C.AxisPositionValues.Left));
                if (needsSerAxis)
                    plotArea.AppendChild(BuildSeriesAxis(serAxisId, valAxisId));
            }
        }

        chart.AppendChild(plotArea);

        ApplyLegendFromProps(chart, properties);

        chart.AppendChild(new C.PlotVisibleOnly { Val = true });
        chart.AppendChild(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });

        // Chart style number — <c:style> precedes <c:chart> in CT_ChartSpace.
        // Drives gridline tint / effect defaults in real PowerPoint; without
        // it a round-tripped chart falls back to the app default style.
        if (properties.TryGetValue("chartStyle", out var chartStyleStr)
            && byte.TryParse(chartStyleStr, out var chartStyleVal)
            && chartStyleVal >= 1 && chartStyleVal <= 48)
        {
            chartSpace.AppendChild(new C.Style { Val = chartStyleVal });
        }

        chartSpace.AppendChild(chart);

        // chartSpace-level default text properties (captured verbatim by the
        // Reader as chartTxPrRaw) — <c:txPr> follows <c:chart> (and c:spPr)
        // in CT_ChartSpace. Sets the base font for every chart element.
        if (properties.TryGetValue("chartTxPrRaw", out var chartTxPrRaw)
            && !string.IsNullOrWhiteSpace(chartTxPrRaw))
        {
            try { chartSpace.AppendChild(new C.TextProperties(chartTxPrRaw)); }
            catch { /* malformed captured XML — keep the chart without it */ }
        }

        // Apply cell references for dotted syntax (series1.values=Sheet1!B2:B13)
        ApplySeriesReferences(plotArea, properties);

        // Restore source c:idx / c:order (series{N}.seriesIdx / .seriesOrder).
        // PowerPoint keys the theme accent cycle and stack order off these —
        // a combo dump reorders series by chart-group, so the positional idx
        // the builders assign recolors every series.
        {
            var allSerForIdx = plotArea.Descendants<OpenXmlCompositeElement>()
                .Where(e => e.LocalName == "ser").ToList();
            for (int i = 0; i < allSerForIdx.Count; i++)
            {
                if (properties.TryGetValue($"series{i + 1}.seriesIdx", out var sIdxStr)
                    && uint.TryParse(sIdxStr, out var sIdxVal))
                {
                    var idxEl = allSerForIdx[i].Elements<C.Index>().FirstOrDefault();
                    if (idxEl != null) idxEl.Val = sIdxVal;
                }
                if (properties.TryGetValue($"series{i + 1}.seriesOrder", out var sOrdStr)
                    && uint.TryParse(sOrdStr, out var sOrdVal))
                {
                    var ordEl = allSerForIdx[i].Elements<C.Order>().FirstOrDefault();
                    if (ordEl != null) ordEl.Val = sOrdVal;
                }
            }
        }

        // Defensive invariant (R26): a built chart must never declare an axis
        // that no chart group references. Orphaned axes (e.g. a secondary
        // axId 3/4 left over from a dump→rebuild with mismatched series
        // binding) make real Excel reject the whole file with 0x800A03EC.
        PruneOrphanAxes(plotArea);

        return chartSpace;
    }

    /// <summary>
    /// Replace literal Values/CategoryAxisData with NumberReference/StringReference
    /// when dotted syntax cell references are used.
    /// </summary>
    private static void ApplySeriesReferences(C.PlotArea plotArea, Dictionary<string, string> properties)
    {
        var extSeries = ParseSeriesDataExtended(properties);
        // Also detect name-only cell references (series{N}.name=Sheet1!A1) so
        // legend text resolves to the cell value instead of a literal string.
        bool hasNameRef = false;
        if (extSeries != null)
        {
            for (int i = 0; i < extSeries.Count; i++)
            {
                if (IsCellReference(extSeries[i].Name)) { hasNameRef = true; break; }
            }
        }
        // R28-B3 — top-level `categories=Sheet1!A1:A3` must rewrite the
        // existing strLit cat to a strRef even when no per-series dotted
        // refs were supplied (extSeries==null). Mirrors R17/R18 series.name
        // and chart-title fixes.
        var topCatRefForBail = ParseCategoriesRef(properties);
        if (extSeries == null || extSeries.Count == 0)
        {
            if (!hasNameRef && topCatRefForBail == null) return;
        }
        if (extSeries != null && !extSeries.Any(s => s.ValuesRef != null || s.CategoriesRef != null || s.BubbleSizeRef != null) && !hasNameRef)
        {
            if (topCatRefForBail == null) return;
        }

        var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();

        // Top-level categories reference applies to all series
        var topCategoriesRef = ParseCategoriesRef(properties);

        // R20-03: when dispBlanksAs=gap, blank source cells must be omitted
        // from the numCache so Excel renders a gap instead of dropping to 0.
        // ParseDataRangeForChart forwards per-series blank index lists in
        // properties[$"series{N}._blankIndexes"] = "1,4,...".
        bool dispBlanksGap = false;
        if (properties.TryGetValue("dispblanksas", out var dba)
            || properties.TryGetValue("dispBlanksAs", out dba)
            || properties.TryGetValue("blanksas", out dba))
        {
            dispBlanksGap = string.Equals(dba?.Trim(), "gap", StringComparison.OrdinalIgnoreCase);
        }

        // R28-B3 — extSeries may be null when the user only set top-level
        // categories=<range> (no series.* dotted keys). Walk all series with
        // an empty info so the topCategoriesRef strRef rewrite still runs.
        int loopCount = extSeries != null ? Math.Min(extSeries.Count, allSer.Count) : allSer.Count;
        for (int i = 0; i < loopCount; i++)
        {
            var info = extSeries != null ? extSeries[i] : new SeriesInfo();
            var ser = allSer[i];

            // Rewrite SeriesText as strRef when the name is a cell reference
            // (e.g. series1.name=Sheet1!A1). Cache is left absent; Excel will
            // resolve the cell on open. See RewriteSeriesTextAsRef for details.
            if (!string.IsNullOrEmpty(info.Name) && IsCellReference(info.Name))
            {
                RewriteSeriesTextAsRef(ser, NormalizeCellReference(info.Name), cachedValue: null);
            }

            // Replace Values (or YValues for scatter/bubble) with NumberReference
            // (preserving literal data as cache).
            if (!string.IsNullOrEmpty(info.ValuesRef))
            {
                HashSet<int>? blanks = null;
                if (dispBlanksGap
                    && properties.TryGetValue($"series{i + 1}._blankIndexes", out var blanksStr)
                    && !string.IsNullOrWhiteSpace(blanksStr))
                {
                    blanks = new HashSet<int>(blanksStr
                        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .Select(s => int.TryParse(s, out var n) ? n : -1)
                        .Where(n => n >= 0));
                }
                // CONSISTENCY(scatter-bubble-no-cat / R21-bt2): scatter and
                // bubble carry y-data in <c:yVal>, not <c:val>.
                OpenXmlCompositeElement? valEl = ser.GetFirstChild<C.Values>()
                    ?? (OpenXmlCompositeElement?)ser.GetFirstChild<C.YValues>();
                if (valEl != null)
                {
                    var numCache = BuildNumberingCacheFromLiteral(
                        valEl.GetFirstChild<C.NumberLiteral>(), blanks);
                    // Source cache formatCode (series{N}.valuesNumFmt, e.g.
                    // #,##0) — sourceLinked data labels render this format.
                    if (numCache != null
                        && properties.TryGetValue($"series{i + 1}.valuesNumFmt", out var vnf)
                        && !string.IsNullOrWhiteSpace(vnf))
                    {
                        var fcEl = numCache.GetFirstChild<C.FormatCode>();
                        if (fcEl != null) fcEl.Text = vnf;
                    }
                    valEl.RemoveAllChildren();
                    var numRef = new C.NumberReference(new C.Formula(info.ValuesRef));
                    if (numCache != null)
                        numRef.AppendChild(numCache);
                    valEl.AppendChild(numRef);
                }
            }

            // Replace CategoryAxisData with StringReference (preserving literal data as cache)
            var catRef = info.CategoriesRef ?? topCategoriesRef;
            if (!string.IsNullOrEmpty(catRef))
            {
                // CONSISTENCY(scatter-bubble-no-cat / R21-bt2): scatter and
                // bubble series use <c:xVal>/<c:yVal>, not <c:cat>/<c:val>.
                // Inserting a <c:cat> on a ScatterChartSeries fails OOXML
                // schema validation (CT_ScatterSer has no cat slot). For these
                // series, rewrite the existing <c:xVal> literal to a number
                // reference so the X axis still tracks the source range.
                bool isScatterOrBubble = ser is C.ScatterChartSeries or C.BubbleChartSeries;
                if (isScatterOrBubble)
                {
                    var xValEl = ser.GetFirstChild<C.XValues>();
                    if (xValEl != null)
                    {
                        var numCache = BuildNumberingCacheFromLiteral(
                            xValEl.GetFirstChild<C.NumberLiteral>(), null);
                        xValEl.RemoveAllChildren();
                        var numRef = new C.NumberReference(new C.Formula(catRef));
                        if (numCache != null)
                            numRef.AppendChild(numCache);
                        xValEl.AppendChild(numRef);
                    }
                    // No cat element to fall back to — for scatter/bubble the
                    // x-data IS the "categories", so silently skip if no xVal.
                }
                else
                {
                    var catEl = ser.GetFirstChild<C.CategoryAxisData>();
                    if (catEl != null)
                    {
                        // A date axis needs NUMERIC categories: rewriting to a
                        // strRef/strCache degrades the dateAx to a text axis in
                        // real PowerPoint (raw serials, no date scaling, no
                        // min/max windowing). When catAxisType=date and every
                        // literal label parses as a number, build numRef +
                        // numCache instead.
                        var catStrLit = catEl.GetFirstChild<C.StringLiteral>();
                        bool dateCats = properties.TryGetValue("catAxisType", out var catTypeStr)
                            && string.Equals(catTypeStr?.Trim(), "date", StringComparison.OrdinalIgnoreCase);
                        if (dateCats && catStrLit != null
                            && catStrLit.Elements<C.StringPoint>().Any()
                            && catStrLit.Elements<C.StringPoint>().All(p =>
                                double.TryParse(p.GetFirstChild<C.NumericValue>()?.Text,
                                    System.Globalization.NumberStyles.Float,
                                    System.Globalization.CultureInfo.InvariantCulture, out _)))
                        {
                            var catNumCache = new C.NumberingCache();
                            catNumCache.AppendChild(new C.FormatCode("General"));
                            var catPtCount = catStrLit.GetFirstChild<C.PointCount>();
                            if (catPtCount != null)
                                catNumCache.AppendChild(new C.PointCount { Val = catPtCount.Val });
                            foreach (var sp in catStrLit.Elements<C.StringPoint>())
                            {
                                var np = new C.NumericPoint { Index = sp.Index };
                                np.AppendChild(new C.NumericValue(sp.GetFirstChild<C.NumericValue>()?.Text ?? ""));
                                catNumCache.AppendChild(np);
                            }
                            catEl.RemoveAllChildren();
                            var catNumRef = new C.NumberReference(new C.Formula(catRef));
                            catNumRef.AppendChild(catNumCache);
                            catEl.AppendChild(catNumRef);
                        }
                        else
                        {
                            var strCache = BuildStringCacheFromLiteral(catStrLit);
                            catEl.RemoveAllChildren();
                            var strRef = new C.StringReference(new C.Formula(catRef));
                            if (strCache != null)
                                strRef.AppendChild(strCache);
                            catEl.AppendChild(strRef);
                        }
                    }
                    else
                    {
                        // Insert CategoryAxisData before Values
                        var valEl = ser.GetFirstChild<C.Values>();
                        var newCat = new C.CategoryAxisData(new C.StringReference(new C.Formula(catRef)));
                        if (valEl != null)
                            valEl.InsertBeforeSelf(newCat);
                        else
                            ser.AppendChild(newCat);
                    }
                }
            }

            // R52 bt-3: rewrite <c:bubbleSize>'s default numLit (BuildBubbleChart
            // seeds size = y-values) to numRef when series{N}.bubbleSize=<range>
            // was supplied. Preserves the cached point values so PowerPoint
            // renders the correct bubble pixel-geometry even before Excel
            // re-evaluates the external workbook reference.
            if (ser is C.BubbleChartSeries && !string.IsNullOrEmpty(info.BubbleSizeRef))
            {
                var bsEl = ser.GetFirstChild<C.BubbleSize>();
                if (bsEl != null)
                {
                    var numCache = BuildNumberingCacheFromLiteral(
                        bsEl.GetFirstChild<C.NumberLiteral>(), null);
                    bsEl.RemoveAllChildren();
                    var numRef = new C.NumberReference(new C.Formula(info.BubbleSizeRef));
                    if (numCache != null)
                        numRef.AppendChild(numCache);
                    bsEl.AppendChild(numRef);
                }
            }
        }
    }

    /// <summary>
    /// Keys that BuildChartSpace doesn't handle directly but SetChartProperties does.
    /// After saving ChartSpace to a ChartPart, call SetChartProperties with these to apply them.
    /// </summary>
    internal static readonly HashSet<string> DeferredAddKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "datalabels", "labels", "labelpos", "labelposition", "labelfont",
        "axistitle", "vtitle", "cattitle", "htitle",
        "axismin", "min", "axismax", "max",
        "majorunit", "minorunit",
        "axisnumfmt", "axisnumberformat",
        "gridlines", "majorgridlines", "minorgridlines",
        // R24 — dotted gridline subkeys (Reader emits gridlineColor /
        // gridlineWidth / gridlineDash; without deferring, Add dropped them).
        "gridlinecolor", "gridlinewidth", "gridlinedash",
        "majorgridlinecolor", "majorgridlinewidth", "majorgridlinedash",
        "minorgridlinecolor", "minorgridlinewidth", "minorgridlinedash",
        "plotareafill", "plotfill", "chartareafill", "chartfill",
        "linewidth", "linedash", "dash", "marker", "markers", "markersize", "markercolor",
        "style", "styleid",
        "transparency", "opacity", "alpha",
        "gradient", "gradients", "gradientfill",
        "trendline",
        "secondaryaxis", "secondary",
        "referenceline", "refline", "targetline",
        "colorrule", "conditionalcolor",
        "combotypes", "combo.types",
        "preset", "style.preset", "theme",
        "view3d", "camera", "perspective",
        "floorraw", "sidewallraw", "backwallraw",
        "holesize", "firstsliceangle", "sliceangle",
        "axisvisible", "axis.visible", "axis.delete",
        "cataxisvisible", "valaxisvisible",
        "majortickmark", "majortick", "minortickmark", "minortick",
        "ticklabelpos", "ticklabelposition",
        "smooth", "showmarker", "showmarkers",
        "scatterstyle", "radarstyle", "varycolors",
        "dispblanksas", "blanksas", "roundedcorners",
        "datatable", "legend.overlay", "legendoverlay",
        "plotarea.border", "plotborder", "chartarea.border", "chartborder",
        "gapwidth", "gap", "overlap",
        "axisline", "axis.line", "cataxisline", "valaxisline",
        // BUG-DUMP-R34-1: verbatim axis-line / plot-area spPr fragments,
        // injected post-build by SetChartProperties at the schema-correct
        // position inside CT_ValAx / CT_CatAx / CT_PlotArea. (plotarea.sppr
        // also matches the "plotarea." deferred prefix; listed for symmetry.)
        "valax.sppr", "catax.sppr", "plotarea.sppr", "chartarea.sppr",
        // verbatim gridline outline spPr (tx1+lumMod tint, cap/cmpd/join) that
        // the granular gridlineColor/Width/Dash keys cannot represent.
        "gridline.sppr", "minorgridline.sppr",
        // BUG-DUMP-R35-1: verbatim per-axis text properties + title paragraph
        // properties, injected post-build at the schema-correct position inside
        // CT_CatAx / CT_ValAx (txPr after spPr, before crossAx) and the title's
        // c:tx/c:rich paragraph (a:pPr). (title.ppr also matches the "title."
        // deferred prefix; listed for symmetry.)
        "valax.txpr", "catax.txpr", "title.ppr",
        // R24 — dotted subkeys mirroring Reader emit.
        "valaxisline.color", "valaxisline.width", "valaxisline.dash",
        "cataxisline.color", "cataxisline.width", "cataxisline.dash",
        "plotarea.border.color", "plotarea.border.width", "plotarea.border.dash",
        "chartarea.border.color", "chartarea.border.width", "chartarea.border.dash",
        "explosion", "explode", "invertifneg", "invertifnegative",
        "errbars", "errorbars", "series.shadow", "seriesshadow",
        "series.outline", "seriesoutline",
        "bubblescale", "shownegbubbles", "sizerepresents",
        "gapdepth", "shape", "barshape", "shape3d",
        "droplines", "hilowlines", "updownbars", "serlines", "serieslines",
        "axisorientation", "axisreverse", "logbase", "logscale", "yaxisscale",
        // CONSISTENCY(cat-axis-type): swap CategoryAxis for DateAxis when
        // catAxisType=date so time-series categories render with date scaling.
        "cataxistype", "categoryaxistype",
        "dispunits", "displayunits", "labeloffset", "ticklabelskip", "tickskip",
        "axisposition", "axispos", "crosses", "crossesat", "crossbetween",
        "plotvisonly", "plotvisibleonly", "autotitledeleted",
        "datalabels.separator", "labelseparator",
        "datalabels.numfmt", "labelnumfmt",
        "datalabels.showleaderlines", "leaderlines", "showleaderlines",
        // CL23 — chart-level trendline.* fan-out
        "trendline.label", "trendline.forecastforward", "trendline.forecastbackward",
        "trendline.order", "trendline.period", "trendline.intercept",
        "trendline.displayequation", "trendline.displayrsquared",
        "errbars.direction", "errbardirection",
        "datalabels.showbubblesize",
        // CleanupE1 — per-flag dotted subkeys for DataLabels on Add.
        "datalabels.showvalue", "datalabels.showval",
        "datalabels.showpercent", "datalabels.showpct",
        "datalabels.showcatname", "datalabels.showcategoryname", "datalabels.showcategory",
        "datalabels.showsername", "datalabels.showseriesname", "datalabels.showseries",
        "datalabels.showlegendkey",
        // R28-B1 — top-level aliases for the dotted datalabels.show* keys above.
        "showvalue", "showval",
        "showpercent", "showpct",
        "showcatname", "showcategoryname", "showcategory",
        "showsername", "showseriesname", "showseries",
        "showlegendkey",
        "axisfont", "axis.font", "legendfont", "legend.font",
        // R15-4: rotate tick labels on cat/val axis. Degrees (e.g. -45).
        "labelrotation", "xaxis.labelrotation", "xaxislabelrotation",
        "valaxis.labelrotation", "valaxislabelrotation", "yaxis.labelrotation", "yaxislabelrotation",
        // Title styling
        "title.font", "titlefont", "title.size", "titlesize",
        "title.color", "titlecolor", "title.bold", "titlebold",
        "title.glow", "titleglow", "title.shadow", "titleshadow",
        // R60 B3 — title.overlay (<c:title><c:overlay val="1"/>) draws the
        // title on top of the plot area. BuildChartTitle defaults to
        // overlay=false; defer so Set can flip it post-build, mirroring
        // legend.overlay.
        "title.overlay", "titleoverlay",
        // Area fill
        "areafill", "area.fill",
        // R24 — bar chart bar-direction (rtl reverses category order on bar
        // groupings). Was previously skipped on Add because it wasn't deferred,
        // so dump→replay dropped the prop.
        "direction",
    };

    /// <summary>
    /// Prefixes for dynamic deferred keys (e.g. title.x, plotArea.y, legend.w,
    /// dataLabel1.text, dataTable.show*, displayUnitsLabel.*, trendlineLabel.*).
    /// </summary>
    private static readonly string[] DeferredPrefixes =
    [
        "title.", "plotarea.", "legend.", "datalabel",
        // legendEntry{N}.delete hides a specific legend entry; the Setter's
        // TryParseLegendEntryKey applies it post-build, so it must defer (the
        // "legend." prefix above does not match "legendentry...").
        "legendentry",
        "datatable.", "displayunitslabel.", "trendlinelabel.",
        "labelfont.",
        // axisTitle.pPr / catTitle.pPr — restore the source axis-title paragraph
        // styling post-build (the axis title element must exist first).
        "axistitle.", "cattitle.",
    ];

    /// <summary>
    /// Check if a property key should be deferred from BuildChartSpace to SetChartProperties.
    /// Matches exact keys in <see cref="DeferredAddKeys"/> plus dynamic prefix patterns.
    /// </summary>
    /// <summary>
    /// Apply the `legend` prop (position, or none via legend=none/false or
    /// hide/hidden) to a chart, inserting the &lt;c:legend&gt; at the correct
    /// CT_Chart position (after plotArea, before plotVisibleOnly/dispBlanksAs).
    /// Shared by the generic build path and the waterfall path (which builds
    /// its own chart and otherwise never consumed `legend`, so legend= was
    /// reported unsupported on waterfall Add).
    /// </summary>
    internal static void ApplyLegendFromProps(C.Chart chart, Dictionary<string, string> properties)
    {
        var showLegend = properties.GetValueOrDefault("legend", "true");
        // CONSISTENCY(legend-hide-alias / R34-1): hide/hidden == legend=none.
        if ((properties.TryGetValue("hide", out var hideVal) && ParseHelpers.IsTruthy(hideVal)) ||
            (properties.TryGetValue("hidden", out var hiddenVal) && ParseHelpers.IsTruthy(hiddenVal)))
            showLegend = "none";
        if (showLegend.Equals("true", StringComparison.OrdinalIgnoreCase))
            showLegend = "bottom";
        if (showLegend.Equals("false", StringComparison.OrdinalIgnoreCase)
            || showLegend.Equals("none", StringComparison.OrdinalIgnoreCase))
            return;
        var legendPos = ParseLegendPosition(showLegend);
        var legend = new C.Legend(
            new C.LegendPosition { Val = legendPos },
            new C.Overlay { Val = false });
        // Schema order: legend precedes plotVisibleOnly/dispBlanksAs. On the
        // generic path those don't exist yet (append == correct); on a
        // pre-built chart (waterfall) insert before them.
        var anchor = (OpenXmlElement?)chart.GetFirstChild<C.PlotVisibleOnly>()
            ?? chart.GetFirstChild<C.DisplayBlanksAs>();
        if (anchor != null) chart.InsertBefore(legend, anchor);
        else chart.AppendChild(legend);
    }

    internal static bool IsDeferredKey(string key)
    {
        if (DeferredAddKeys.Contains(key)) return true;
        var lower = key.ToLowerInvariant();
        foreach (var prefix in DeferredPrefixes)
            if (lower.StartsWith(prefix)) return true;
        // CONSISTENCY(chart-series-color): select per-series dotted keys
        // route through HandleSeriesDottedProperty at SetChartProperties
        // time. Only visual-effect subkeys are deferred here; `.name`,
        // `.values`, `.categories`, `.ref`, `.valuesRef`, `.categoriesRef`,
        // `.color` are consumed at build time by ParseSeriesData /
        // ParseSeriesColors and must NOT be deferred (double-apply /
        // literal-expansion regressions).
        if (TryParseSeriesDottedKey(key, out _, out var sProp))
        {
            if (DeferredSeriesSubkeys.Contains(sProp)) return true;
            // R38: per-point sub-keys (point{M}.color / .explosion / .marker
            // / .markerSize / .markerColor) dispatch via the same
            // HandleSeriesDottedProperty path but the sub-key carries the
            // M index so an exact-set membership check misses them.
            if (System.Text.RegularExpressions.Regex.IsMatch(
                    sProp, @"^point\d+\.", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                return true;
        }
        return false;
    }

    // Per-series dotted subkeys that route through HandleSeriesDottedProperty
    // during SetChartProperties (post-build). See IsDeferredKey.
    private static readonly HashSet<string> DeferredSeriesSubkeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "gradient", "gradientfill",
        // BUG-DUMP-R33-1: verbatim styling fragments captured by the Reader
        // (series <c:spPr>, the per-data-point <c:dPt> list, and the rich
        // <c:dLbls>) are injected post-build by HandleSeriesDottedProperty.
        "sppr", "dpt", "dlbls",
        "smooth", "trendline", "marker", "markersize", "markercolor",
        "invertifneg", "invertifnegative",
        "errbars", "errorbars",
        "explosion", "explode",
        "linewidth", "linedash", "dash",
        "shadow", "outline",
        "outlinecolor", "outlinewidth", "outlinedash",
        "alpha", "transparency",
        // R38: per-series labelFont subkeys — Setter dispatches via
        // HandleSeriesDottedProperty (default arm `labelfont` / `labelfont.*`).
        // Without deferral, AddChart's build-time parse drops them and
        // dump→replay loses series-scoped data label fonts.
        "labelfont",
        "labelfont.color", "labelfont.size", "labelfont.bold",
        "labelfont.name", "labelfont.font",
    };

    // ==================== Chart Type Builders ====================

    /// <summary>
    /// Build one homogeneous CT_*Chart container for a contiguous slice of
    /// the combo's seriesData and append it to plotArea. Series indices remain
    /// global (startIdx..endIdxExclusive-1) so each ser's c:idx matches its
    /// original position — required for the Reader's comboTypes round-trip.
    /// </summary>
    private static void BuildComboGroup(string typeLabel,
        C.PlotArea plotArea,
        List<(string name, double[] values)> seriesData,
        string[]? categories,
        int startIdx, int endIdxExclusive,
        uint catAxisId, uint valAxisId,
        string[]? colors,
        HashSet<int>? noFillSeries = null)
    {
        // Grouping-qualified tokens (columnstacked / areapercentstacked …)
        // from the Reader's comboTypes emit — parse the suffix so a stacked
        // combo group doesn't rebuild as clustered/standard.
        string grpSuffix = "";
        if (typeLabel.EndsWith("percentstacked", StringComparison.Ordinal))
        { grpSuffix = "percentstacked"; typeLabel = typeLabel[..^14]; }
        else if (typeLabel.EndsWith("stacked", StringComparison.Ordinal))
        { grpSuffix = "stacked"; typeLabel = typeLabel[..^7]; }
        var barGrp = grpSuffix switch
        {
            "percentstacked" => C.BarGroupingValues.PercentStacked,
            "stacked" => C.BarGroupingValues.Stacked,
            _ => C.BarGroupingValues.Clustered,
        };
        var stdGrp = grpSuffix switch
        {
            "percentstacked" => C.GroupingValues.PercentStacked,
            "stacked" => C.GroupingValues.Stacked,
            _ => C.GroupingValues.Standard,
        };
        OpenXmlCompositeElement container = typeLabel switch
        {
            "bar" => new C.BarChart(
                new C.BarDirection { Val = C.BarDirectionValues.Bar },
                new C.BarGrouping { Val = barGrp },
                new C.VaryColors { Val = false }),
            "column" => new C.BarChart(
                new C.BarDirection { Val = C.BarDirectionValues.Column },
                new C.BarGrouping { Val = barGrp },
                new C.VaryColors { Val = false }),
            "area" => new C.AreaChart(
                new C.Grouping { Val = stdGrp },
                new C.VaryColors { Val = false }),
            "line" => new C.LineChart(
                new C.Grouping { Val = stdGrp },
                new C.VaryColors { Val = false }),
            _ => new C.LineChart(
                new C.Grouping { Val = stdGrp },
                new C.VaryColors { Val = false }),
        };

        for (int i = startIdx; i < endIdxExclusive; i++)
        {
            var clr = ResolveSeriesColor(colors, i, noFillSeries);
            OpenXmlCompositeElement ser = typeLabel switch
            {
                "bar" or "column" => BuildBarSeries((uint)i, seriesData[i].name,
                    categories, seriesData[i].values, clr),
                "area" => BuildAreaSeries((uint)i, seriesData[i].name,
                    categories, seriesData[i].values, clr),
                _ => BuildLineSeries((uint)i, seriesData[i].name,
                    categories, seriesData[i].values, clr),
            };
            container.AppendChild(ser);
        }

        if (typeLabel == "bar" || typeLabel == "column")
            container.AppendChild(new C.GapWidth { Val = 150 });

        container.AppendChild(new C.AxisId { Val = catAxisId });
        container.AppendChild(new C.AxisId { Val = valAxisId });
        plotArea.AppendChild(container);
    }

    internal static C.BarChart BuildBarChart(
        C.BarDirectionValues direction, bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null,
        HashSet<int>? noFillSeries = null)
    {
        var grouping = percentStacked ? C.BarGroupingValues.PercentStacked
            : stacked ? C.BarGroupingValues.Stacked
            : C.BarGroupingValues.Clustered;

        var barChart = new C.BarChart(
            new C.BarDirection { Val = direction },
            new C.BarGrouping { Val = grouping },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = ResolveSeriesColor(colors, i, noFillSeries);
            barChart.AppendChild(BuildBarSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values, color));
        }

        barChart.AppendChild(new C.GapWidth { Val = (ushort)150 });
        if (stacked || percentStacked)
            barChart.AppendChild(new C.Overlap { Val = 100 });
        barChart.AppendChild(new C.AxisId { Val = catAxisId });
        barChart.AppendChild(new C.AxisId { Val = valAxisId });
        return barChart;
    }

    internal static C.LineChart BuildLineChart(
        bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null,
        HashSet<int>? noFillSeries = null)
    {
        var grouping = percentStacked ? C.GroupingValues.PercentStacked
            : stacked ? C.GroupingValues.Stacked
            : C.GroupingValues.Standard;

        var lineChart = new C.LineChart(
            new C.Grouping { Val = grouping },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = ResolveSeriesColor(colors, i, noFillSeries);
            lineChart.AppendChild(BuildLineSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values, color));
        }

        // ShowMarker is opt-in: only emit when user sets the prop (Setter
        // handles "showmarker" / "showmarkers"). Hard-coding true broke
        // dump→replay for line charts that should have no markers.
        lineChart.AppendChild(new C.AxisId { Val = catAxisId });
        lineChart.AppendChild(new C.AxisId { Val = valAxisId });
        return lineChart;
    }

    internal static C.AreaChart BuildAreaChart(
        bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null,
        HashSet<int>? noFillSeries = null)
    {
        var grouping = percentStacked ? C.GroupingValues.PercentStacked
            : stacked ? C.GroupingValues.Stacked
            : C.GroupingValues.Standard;

        var areaChart = new C.AreaChart(
            new C.Grouping { Val = grouping },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = ResolveSeriesColor(colors, i, noFillSeries);
            areaChart.AppendChild(BuildAreaSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values, color));
        }

        areaChart.AppendChild(new C.AxisId { Val = catAxisId });
        areaChart.AppendChild(new C.AxisId { Val = valAxisId });
        return areaChart;
    }

    internal static C.PieChart BuildPieChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        string[]? colors = null)
    {
        var pieChart = new C.PieChart(new C.VaryColors { Val = true });
        if (seriesData.Count > 0)
        {
            var series = BuildPieSeries(0, seriesData[0].name,
                categories, seriesData[0].values);
            ApplyDataPointColors(series, seriesData[0].values.Length, colors);
            pieChart.AppendChild(series);
        }
        return pieChart;
    }

    /// <summary>
    /// Build a c:ofPieChart (pieOfPie / barOfPie) — splits the trailing
    /// data points of a single pie series into a secondary pie/bar so the
    /// small-value slices remain readable. CT_OfPieChart requires an
    /// ofPieType ("pie" or "bar") plus exactly one c:ser; the SecondPiePoints
    /// knob defaults to 3 (matching Excel's default split).
    /// </summary>
    internal static C.OfPieChart BuildOfPieChart(
        C.OfPieValues ofPieType,
        string[]? categories, List<(string name, double[] values)> seriesData,
        string[]? colors = null)
    {
        var ofPieChart = new C.OfPieChart(
            new C.OfPieType { Val = ofPieType },
            new C.VaryColors { Val = true }
        );
        if (seriesData.Count > 0)
        {
            var series = BuildPieSeries(0, seriesData[0].name,
                categories, seriesData[0].values);
            ApplyDataPointColors(series, seriesData[0].values.Length, colors);
            ofPieChart.AppendChild(series);
        }
        // Default split = 3 trailing points (Excel's default). Document syntax:
        // user can later set via secondPieSize / splitPos in a future round.
        ofPieChart.AppendChild(new C.SecondPieSize { Val = 75 });
        ofPieChart.AppendChild(new C.SeriesLines());
        return ofPieChart;
    }

    internal static C.DoughnutChart BuildDoughnutChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        string[]? colors = null)
    {
        var chart = new C.DoughnutChart(new C.VaryColors { Val = true });
        // Unlike pie, a doughnut legitimately supports multiple series —
        // PowerPoint renders one concentric ring per series. Emit one
        // <c:ser> per series (was previously hardcoded to seriesData[0],
        // silently dropping inner rings). Per-data-point colors apply to
        // each series' slices (PowerPoint colors every ring by category).
        for (int s = 0; s < seriesData.Count; s++)
        {
            var series = BuildPieSeries((uint)s, seriesData[s].name,
                categories, seriesData[s].values);
            ApplyDataPointColors(series, seriesData[s].values.Length, colors);
            chart.AppendChild(series);
        }
        chart.AppendChild(new C.HoleSize { Val = 50 });
        return chart;
    }

    /// <summary>
    /// For pie/doughnut charts, apply per-data-point colors via c:dPt elements.
    /// Each slice gets its own DataPoint with Index and ChartShapeProperties containing a solid fill.
    /// </summary>
    private static void ApplyDataPointColors(C.PieChartSeries series, int pointCount, string[]? colors)
    {
        if (colors == null || colors.Length == 0) return;
        var count = Math.Min(pointCount, colors.Length);
        for (int i = 0; i < count; i++)
        {
            ApplyDataPointColor(series, i, colors[i]);
        }
    }

    internal static C.ScatterChart BuildScatterChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string scatterStyle = "lineMarker",
        string[]? colors = null)
    {
        var styleVal = scatterStyle.ToLowerInvariant() switch
        {
            "marker" => C.ScatterStyleValues.Marker,
            "line" => C.ScatterStyleValues.Line,
            "smooth" => C.ScatterStyleValues.SmoothMarker,
            "smoothmarker" => C.ScatterStyleValues.SmoothMarker,
            _ => C.ScatterStyleValues.LineMarker
        };
        var scatterChart = new C.ScatterChart(
            new C.ScatterStyle { Val = styleVal },
            new C.VaryColors { Val = false }
        );

        double[]? xValues = null;
        if (categories != null)
            xValues = categories.Select(c => double.TryParse(c, out var v) ? v : 0).ToArray();

        var hideLines = styleVal == C.ScatterStyleValues.Marker;
        for (int i = 0; i < seriesData.Count; i++)
        {
            var ser = BuildScatterSeries((uint)i, seriesData[i].name,
                xValues, seriesData[i].values);
            // For marker-only style, explicitly hide connecting lines.
            // CT_ScatterSer schema: idx, order, tx, spPr, marker, dPt*, dLbls?,
            // trendline*, errBars?, xVal?, yVal?, smooth?, extLst? — spPr must
            // sit right after tx; appending at end loses it on strict readers.
            // CONSISTENCY(chart-schema-order): route through the same helper used
            // by per-series fill/line/effect setters.
            if (hideLines)
            {
                var spPr = GetOrCreateSeriesShapeProperties(ser);
                spPr.RemoveAllChildren<Drawing.Outline>();
                spPr.AppendChild(new Drawing.Outline(new Drawing.NoFill()));
            }
            // CONSISTENCY(series-color-builder): mirror BuildBarChart /
            // BuildLineChart — when user supplied colors[] via series{N}.color
            // or a colors property, apply per-series so scatter rounds-trip
            // a non-default palette through dump→replay.
            if (colors != null && i < colors.Length && !string.IsNullOrEmpty(colors[i]))
                ApplySeriesColor(ser, colors[i]);
            scatterChart.AppendChild(ser);
        }

        scatterChart.AppendChild(new C.AxisId { Val = catAxisId });
        scatterChart.AppendChild(new C.AxisId { Val = valAxisId });
        return scatterChart;
    }

    // ==================== Bubble Chart ====================

    internal static C.BubbleChart BuildBubbleChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null,
        Dictionary<string, string>? properties = null)
    {
        var bubbleChart = new C.BubbleChart(new C.VaryColors { Val = false });

        // R16-7: per-series literal bubble sizes (series{N}.bubbleSize=...). The
        // default path below seeds size = y-values; consume the user-supplied
        // literal sizes here so they're written verbatim, mirroring how the
        // ref form (series{N}.bubbleSizeRef) is applied in PostProcessSeriesRefs.
        var extSeries = properties != null ? ParseSeriesDataExtended(properties) : null;

        double[]? xValues = null;
        if (categories != null)
            xValues = categories.Select(c => double.TryParse(c, out var v) ? v : 0).ToArray();

        // Bubble-specific data shape: data="X:..;Y:..;Size:.." describes ONE
        // bubble series with X/Y/Size triples — not three parallel categorical
        // series (the old behaviour stacked all bubbles at the last X value
        // because each "series" emitted a y-line at the same x indices).
        // Detect the X/Y/Size sentinel (case-insensitive) and collapse to a
        // single C.BubbleChartSeries. Falls back to the per-series path
        // (legacy: each row is its own bubble series) when the names don't
        // match the triple.
        if (seriesData.Count >= 2)
        {
            var byName = seriesData.ToDictionary(
                s => s.name.Trim().ToLowerInvariant(),
                s => s.values);
            if (byName.ContainsKey("x") && byName.ContainsKey("y"))
            {
                var xs = byName["x"];
                var ys = byName["y"];
                var sizes = byName.TryGetValue("size", out var sz) ? sz
                          : byName.TryGetValue("r", out var rsz) ? rsz
                          : ys;
                var color = colors != null && colors.Length > 0 ? colors[0] : DefaultSeriesColors[0];
                var series = new C.BubbleChartSeries(
                    new C.Index { Val = 0 },
                    new C.Order { Val = 0 },
                    new C.SeriesText(new C.NumericValue("Bubble"))
                );
                ApplySeriesColor(series, color);

                var xLit = new C.NumberLiteral(new C.PointCount { Val = (uint)xs.Length });
                for (int j = 0; j < xs.Length; j++)
                    xLit.AppendChild(new C.NumericPoint(new C.NumericValue(xs[j].ToString("G"))) { Index = (uint)j });
                series.AppendChild(new C.XValues(xLit));

                var yLit = new C.NumberLiteral(new C.PointCount { Val = (uint)ys.Length });
                for (int j = 0; j < ys.Length; j++)
                    yLit.AppendChild(new C.NumericPoint(new C.NumericValue(ys[j].ToString("G"))) { Index = (uint)j });
                series.AppendChild(new C.YValues(yLit));

                var sizeLit = new C.NumberLiteral(new C.PointCount { Val = (uint)sizes.Length });
                for (int j = 0; j < sizes.Length; j++)
                    sizeLit.AppendChild(new C.NumericPoint(new C.NumericValue(sizes[j].ToString("G"))) { Index = (uint)j });
                series.AppendChild(new C.BubbleSize(sizeLit));

                bubbleChart.AppendChild(series);
                bubbleChart.AppendChild(new C.AxisId { Val = catAxisId });
                bubbleChart.AppendChild(new C.AxisId { Val = valAxisId });
                return bubbleChart;
            }
        }

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            var (name, values) = seriesData[i];
            var series = new C.BubbleChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                BuildSeriesText(name)
            );
            ApplySeriesColor(series, color);

            if (xValues != null)
            {
                var xLit = new C.NumberLiteral(new C.PointCount { Val = (uint)xValues.Length });
                for (int j = 0; j < xValues.Length; j++)
                    xLit.AppendChild(new C.NumericPoint(new C.NumericValue(xValues[j].ToString("G"))) { Index = (uint)j });
                series.AppendChild(new C.XValues(xLit));
            }

            var yLit = new C.NumberLiteral(new C.PointCount { Val = (uint)values.Length });
            for (int j = 0; j < values.Length; j++)
                yLit.AppendChild(new C.NumericPoint(new C.NumericValue(values[j].ToString("G"))) { Index = (uint)j });
            series.AppendChild(new C.YValues(yLit));

            // Bubble sizes — user-supplied literal sizes (R16-7) when present,
            // otherwise default to the y-values.
            var sizes = (extSeries != null && i < extSeries.Count && extSeries[i].BubbleSizeValues != null)
                ? extSeries[i].BubbleSizeValues!
                : values;
            var sizeLit = new C.NumberLiteral(new C.PointCount { Val = (uint)sizes.Length });
            for (int j = 0; j < sizes.Length; j++)
                sizeLit.AppendChild(new C.NumericPoint(new C.NumericValue(sizes[j].ToString("G"))) { Index = (uint)j });
            series.AppendChild(new C.BubbleSize(sizeLit));

            bubbleChart.AppendChild(series);
        }

        bubbleChart.AppendChild(new C.AxisId { Val = catAxisId });
        bubbleChart.AppendChild(new C.AxisId { Val = valAxisId });
        return bubbleChart;
    }

    // ==================== Radar Chart ====================

    internal static C.RadarChart BuildRadarChart(
        string radarStyle,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null)
    {
        var style = radarStyle.ToLowerInvariant() switch
        {
            "filled" or "fill" => C.RadarStyleValues.Filled,
            "marker" => C.RadarStyleValues.Marker,
            _ => C.RadarStyleValues.Standard
        };

        var radarChart = new C.RadarChart(
            new C.RadarStyle { Val = style },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            var series = new C.RadarChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                BuildSeriesText(seriesData[i].name)
            );
            ApplySeriesColor(series, color);
            if (categories != null) series.AppendChild(BuildCategoryData(categories));
            series.AppendChild(BuildValues(seriesData[i].values));
            radarChart.AppendChild(series);
        }

        radarChart.AppendChild(new C.AxisId { Val = catAxisId });
        radarChart.AppendChild(new C.AxisId { Val = valAxisId });
        return radarChart;
    }

    // ==================== Stock Chart ====================

    internal static C.StockChart BuildStockChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId)
    {
        // Stock chart expects series in Open-High-Low-Close order (4 series)
        // or High-Low-Close order (3 series)
        // BUG-R6A(BUG2): OOXML CT_StockChart constrains c:ser to minOccurs=3,
        // maxOccurs=4. Any other count writes a schema-invalid file (e.g.
        // c:hiLowLines appears where the validator still expects c:ser). Reject
        // up front with an actionable message rather than emitting broken XML.
        if (seriesData.Count is < 3 or > 4)
            throw new ArgumentException(
                $"stock chart requires 3 or 4 data series (high, low, close [, open]); got {seriesData.Count}.");

        var stockChart = new C.StockChart();

        for (int i = 0; i < seriesData.Count; i++)
        {
            var series = new C.LineChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                BuildSeriesText(seriesData[i].name)
            );

            // Hide individual series lines — stock chart visuals come from
            // hiLowLines + upDownBars, not from the series lines themselves
            var spPr = new C.ChartShapeProperties();
            spPr.AppendChild(new Drawing.Outline(new Drawing.NoFill()));
            series.AppendChild(spPr);

            // No markers on stock series
            series.AppendChild(new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.None }));

            if (categories != null) series.AppendChild(BuildCategoryData(categories));
            series.AppendChild(BuildValues(seriesData[i].values));
            stockChart.AppendChild(series);
        }

        // Hi-low lines: vertical lines connecting High to Low at each data point
        stockChart.AppendChild(new C.HighLowLines());

        // Up-down bars: colored boxes from Open to Close (green=up, red=down)
        if (seriesData.Count >= 4)
        {
            var upDownBars = new C.UpDownBars(
                new C.GapWidth { Val = 150 }
            );
            var upBars = new C.UpBars();
            var upSpPr = new C.ChartShapeProperties();
            upSpPr.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "4CAF50" }));
            upBars.AppendChild(upSpPr);
            upDownBars.AppendChild(upBars);
            var downBars = new C.DownBars();
            var dnSpPr = new C.ChartShapeProperties();
            dnSpPr.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "F44336" }));
            downBars.AppendChild(dnSpPr);
            upDownBars.AppendChild(downBars);
            stockChart.AppendChild(upDownBars);
        }

        stockChart.AppendChild(new C.AxisId { Val = catAxisId });
        stockChart.AppendChild(new C.AxisId { Val = valAxisId });
        return stockChart;
    }

    // ==================== Default Series Colors ====================

    // CONSISTENCY(chart-default-palette): canonical source is
    // OfficeDefaultThemeColors.DefaultChartSeriesPalette so the OOXML
    // builder and the SVG preview renderer cannot drift apart.
    internal static readonly string[] DefaultSeriesColors =
        OfficeDefaultThemeColors.DefaultChartSeriesPalette;

    // BUG-DUMP-R36-3: resolve the fill color for series i. An explicit color
    // (positional `colors=` or `series{N}.color`) wins. Otherwise the default
    // accent palette is used — EXCEPT for series the dump flagged as carrying
    // per-point dPt styling with no series-level fill, which return null so
    // BuildBarSeries/BuildLineSeries/... emit no series <c:spPr> (preserving the
    // source's "no series fill, rely on dPt + varyColors" shape).
    private static string? ResolveSeriesColor(string[]? colors, int i, HashSet<int>? noFillSeries)
    {
        if (colors != null && i < colors.Length && !string.IsNullOrEmpty(colors[i]))
            return colors[i];
        if (noFillSeries != null && noFillSeries.Contains(i))
            return null;
        return DefaultSeriesColors[i % DefaultSeriesColors.Length];
    }

    // ==================== Series Color ====================

    /// <summary>
    /// Apply per-series dotted colors (`series{N}.color=<hex>`) to the
    /// provided series elements. Used by single-series chart builders
    /// (pie/doughnut/stock) where positional `colors=` is per-data-point.
    /// </summary>
    internal static void ApplyDottedSeriesColors(OpenXmlCompositeElement[] seriesElems, Dictionary<int, string> dotted)
    {
        if (dotted == null || dotted.Count == 0) return;
        for (int i = 0; i < seriesElems.Length; i++)
        {
            if (dotted.TryGetValue(i + 1, out var c) && !string.IsNullOrEmpty(c))
                ApplySeriesColor(seriesElems[i], c);
        }
    }

    internal static void ApplySeriesColor(OpenXmlCompositeElement series, string color)
    {
        series.RemoveAllChildren<C.ChartShapeProperties>();
        var spPr = new C.ChartShapeProperties();
        if (color.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            spPr.AppendChild(new Drawing.NoFill());
            var serText0 = series.GetFirstChild<C.SeriesText>();
            if (serText0 != null) serText0.InsertAfterSelf(spPr);
            else series.PrependChild(spPr);
            return;
        }
        // Line-based series (line/scatter/radar) carry their color on the
        // line stroke (<a:ln><a:solidFill>); real PowerPoint IGNORES a bare
        // <a:solidFill> for the stroke and falls back to the theme color.
        // Area-based series (bar/column/pie/area/doughnut) use the bare
        // <a:solidFill> as an area fill. All series share LocalName "ser", so
        // detect by the strongly-typed element class (works pre-append when
        // Parent is still null) and fall back to the parent chart type for
        // loosely-typed elements obtained from an already-parsed tree.
        bool isLineBased = series is C.LineChartSeries or C.ScatterChartSeries or C.RadarChartSeries
            || series.Parent?.LocalName is "lineChart" or "scatterChart" or "radarChart";

        if (isLineBased)
        {
            const int defaultStrokeWidthEmu = 25400; // 2pt × 12700 EMU/pt
            var outline = new Drawing.Outline { Width = defaultStrokeWidthEmu };
            var lnFill = new Drawing.SolidFill();
            lnFill.AppendChild(BuildChartColorElement(color));
            outline.AppendChild(lnFill);
            spPr.AppendChild(outline);
        }
        else
        {
            var solidFill = new Drawing.SolidFill();
            solidFill.AppendChild(BuildChartColorElement(color));
            spPr.AppendChild(solidFill);
        }

        var serText = series.GetFirstChild<C.SeriesText>();
        if (serText != null)
            serText.InsertAfterSelf(spPr);
        else
            series.PrependChild(spPr);
    }

    /// <summary>
    /// Build a fill element: solid if single color, gradient if contains '-'.
    /// Gradient format: "color1-color2[:angle]" or "color1-color2-color3[:angle]"
    /// Internal so the cx (ExtendedChart) builder can share the exact same fill
    /// vocabulary (solid / gradient / pattern / none) as regular cCharts.
    /// </summary>
    internal static OpenXmlElement BuildFillElement(string value)
    {
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
            return new Drawing.NoFill();

        // Pattern fill: "pattern:preset[:fg[:bg]]" — mirrors the spec form
        // emitted by ReadPatternSpec on dump. Bare "pattern" (preset lost on
        // read) and bare "blip" still degrade to solid black so unknown hints
        // don't crash batch replay (preserves R62 5b91dfbe "key survives" goal).
        if (value.StartsWith("pattern:", StringComparison.OrdinalIgnoreCase))
            return BuildChartPatternFill(value.Substring("pattern:".Length));
        if (value.Equals("pattern", StringComparison.OrdinalIgnoreCase)
            || value.Equals("blip", StringComparison.OrdinalIgnoreCase))
        {
            var fallback = new Drawing.SolidFill();
            fallback.AppendChild(BuildChartColorElement("000000"));
            return fallback;
        }

        // R63 t-2: pre-strip an optional trailing ":sN" scaled marker so the
        // existing LastIndexOf(':') angle parser keeps working. Reader emits
        // ":s0" only for <a:lin scaled="0"> sources; absent means scaled=true.
        bool scaledFlag = true;
        if (value.EndsWith(":s0", StringComparison.Ordinal))
        {
            scaledFlag = false;
            value = value[..^3];
        }
        else if (value.EndsWith(":s1", StringComparison.Ordinal))
        {
            value = value[..^3];
        }

        // Check if it's a gradient (contains - but not a single hex with alpha like 80FF0000)
        var colonIdx = value.LastIndexOf(':');
        var colorPart = colonIdx > 6 ? value[..colonIdx] : value;
        if (colorPart.Contains('-') && colorPart.Split('-').Length >= 2 && colorPart.Split('-')[0].Length <= 8)
        {
            // Gradient: reuse ApplySeriesGradient logic
            var anglePart = 0;
            if (colonIdx > 0 && int.TryParse(value[(colonIdx + 1)..], out var angle))
                anglePart = angle;
            else
                colonIdx = -1;

            var colors = (colonIdx > 0 ? value[..colonIdx] : value).Split('-').Select(c => c.Trim()).ToArray();
            var gradFill = new Drawing.GradientFill();
            var gsLst = new Drawing.GradientStopList();
            for (int i = 0; i < colors.Length; i++)
            {
                var pos = colors.Length == 1 ? 0 : (int)(i * 100000.0 / (colors.Length - 1));
                var gs = new Drawing.GradientStop { Position = pos };
                gs.AppendChild(BuildChartColorElement(colors[i]));
                gsLst.AppendChild(gs);
            }
            gradFill.AppendChild(gsLst);
            gradFill.AppendChild(new Drawing.LinearGradientFill { Angle = ParseHelpers.GradientAngleToOoxmlUnits(anglePart), Scaled = scaledFlag });
            return gradFill;
        }

        // Solid fill
        var solidFill = new Drawing.SolidFill();
        solidFill.AppendChild(BuildChartColorElement(value));
        return solidFill;
    }

    /// <summary>
    /// Build a chart-area / plot-area a:pattFill from the compound spec body
    /// "preset[:fg[:bg]]" (the "pattern:" prefix is stripped by the caller).
    /// Mirrors PowerPointHandler.Fill.BuildPatternFill semantics: empty fg
    /// defaults to 000000, empty bg defaults to FFFFFF. Unknown preset names
    /// silently fall through to Drawing.PresetPatternValues.Percent50 — same
    /// "key survives, looks vaguely right" graceful-degradation contract as
    /// the bare "pattern" / "blip" fallback above. Schema order: fgClr → bgClr.
    /// </summary>
    private static Drawing.PatternFill BuildChartPatternFill(string body)
    {
        var parts = body.Split(':');
        var presetName = parts.Length > 0 ? parts[0].Trim() : "";
        var fg = parts.Length > 1 && !string.IsNullOrWhiteSpace(parts[1]) ? parts[1].Trim() : "000000";
        var bg = parts.Length > 2 && !string.IsNullOrWhiteSpace(parts[2]) ? parts[2].Trim() : "FFFFFF";

        var patternFill = new Drawing.PatternFill { Preset = ParseChartPresetPattern(presetName) };
        var fgClr = new Drawing.ForegroundColor();
        fgClr.AppendChild(BuildChartColorElement(fg));
        patternFill.AppendChild(fgClr);
        var bgClr = new Drawing.BackgroundColor();
        bgClr.AppendChild(BuildChartColorElement(bg));
        patternFill.AppendChild(bgClr);
        return patternFill;
    }

    /// <summary>
    /// Parse an a:pattFill preset attribute (e.g. "pct50", "ltHorz", "cross").
    /// Case-insensitive lookup over the OOXML preset names. Unknown names
    /// default to Percent50 — same graceful-degradation policy as the
    /// "pattern" / "blip" bare-hint fallback. Mirrors ParsePresetPattern in
    /// PowerPointHandler.Fill.cs (intentional duplicate to keep ChartHelper
    /// self-contained — chart fill builders do not depend on the pptx handler).
    /// </summary>
    private static readonly Dictionary<string, Drawing.PresetPatternValues> _chartPresetPatternMap
        = new(StringComparer.OrdinalIgnoreCase)
        {
            ["pct5"] = Drawing.PresetPatternValues.Percent5,
            ["pct10"] = Drawing.PresetPatternValues.Percent10,
            ["pct20"] = Drawing.PresetPatternValues.Percent20,
            ["pct25"] = Drawing.PresetPatternValues.Percent25,
            ["pct30"] = Drawing.PresetPatternValues.Percent30,
            ["pct40"] = Drawing.PresetPatternValues.Percent40,
            ["pct50"] = Drawing.PresetPatternValues.Percent50,
            ["pct60"] = Drawing.PresetPatternValues.Percent60,
            ["pct70"] = Drawing.PresetPatternValues.Percent70,
            ["pct75"] = Drawing.PresetPatternValues.Percent75,
            ["pct80"] = Drawing.PresetPatternValues.Percent80,
            ["pct90"] = Drawing.PresetPatternValues.Percent90,
            ["horz"] = Drawing.PresetPatternValues.Horizontal,
            ["vert"] = Drawing.PresetPatternValues.Vertical,
            ["ltHorz"] = Drawing.PresetPatternValues.LightHorizontal,
            ["ltVert"] = Drawing.PresetPatternValues.LightVertical,
            ["ltDnDiag"] = Drawing.PresetPatternValues.LightDownwardDiagonal,
            ["ltUpDiag"] = Drawing.PresetPatternValues.LightUpwardDiagonal,
            ["dkHorz"] = Drawing.PresetPatternValues.DarkHorizontal,
            ["dkVert"] = Drawing.PresetPatternValues.DarkVertical,
            ["dkDnDiag"] = Drawing.PresetPatternValues.DarkDownwardDiagonal,
            ["dkUpDiag"] = Drawing.PresetPatternValues.DarkUpwardDiagonal,
            ["narHorz"] = Drawing.PresetPatternValues.NarrowHorizontal,
            ["narVert"] = Drawing.PresetPatternValues.NarrowVertical,
            ["dnDiag"] = Drawing.PresetPatternValues.DownwardDiagonal,
            ["upDiag"] = Drawing.PresetPatternValues.UpwardDiagonal,
            ["wdUpDiag"] = Drawing.PresetPatternValues.WideUpwardDiagonal,
            ["wdDnDiag"] = Drawing.PresetPatternValues.WideDownwardDiagonal,
            ["dashHorz"] = Drawing.PresetPatternValues.DashedHorizontal,
            ["dashVert"] = Drawing.PresetPatternValues.DashedVertical,
            ["dashDnDiag"] = Drawing.PresetPatternValues.DashedDownwardDiagonal,
            ["dashUpDiag"] = Drawing.PresetPatternValues.DashedUpwardDiagonal,
            ["cross"] = Drawing.PresetPatternValues.Cross,
            ["diagCross"] = Drawing.PresetPatternValues.DiagonalCross,
            ["smGrid"] = Drawing.PresetPatternValues.SmallGrid,
            ["lgGrid"] = Drawing.PresetPatternValues.LargeGrid,
            ["smConfetti"] = Drawing.PresetPatternValues.SmallConfetti,
            ["lgConfetti"] = Drawing.PresetPatternValues.LargeConfetti,
            ["horzBrick"] = Drawing.PresetPatternValues.HorizontalBrick,
            ["diagBrick"] = Drawing.PresetPatternValues.DiagonalBrick,
            ["solidDmnd"] = Drawing.PresetPatternValues.SolidDiamond,
            ["openDmnd"] = Drawing.PresetPatternValues.OpenDiamond,
            ["dotDmnd"] = Drawing.PresetPatternValues.DottedDiamond,
            ["plaid"] = Drawing.PresetPatternValues.Plaid,
            ["sphere"] = Drawing.PresetPatternValues.Sphere,
            ["weave"] = Drawing.PresetPatternValues.Weave,
            ["divot"] = Drawing.PresetPatternValues.Divot,
            ["shingle"] = Drawing.PresetPatternValues.Shingle,
            ["wave"] = Drawing.PresetPatternValues.Wave,
            ["trellis"] = Drawing.PresetPatternValues.Trellis,
            ["zigZag"] = Drawing.PresetPatternValues.ZigZag,
        };

    private static Drawing.PresetPatternValues ParseChartPresetPattern(string name)
    {
        // Case-insensitive lookup by camelCase XML token (what ReadPatternSpec
        // emits via Preset.InnerText). Also accept the SDK enum member name
        // (e.g. "Percent50") so anything Enum.ToString() could have produced
        // is round-trip safe. Unknown presets fall back to Percent50 — same
        // "key survives" graceful-degradation as the bare "pattern" hint.
        if (_chartPresetPatternMap.TryGetValue(name, out var v)) return v;
        foreach (var kv in _chartPresetPatternMap)
        {
            if (string.Equals(kv.Value.ToString(), name, StringComparison.OrdinalIgnoreCase))
                return kv.Value;
        }
        return Drawing.PresetPatternValues.Percent50;
    }

    /// <summary>
    /// Parse a compound text-style spec "size:color:fontname" into a
    /// <see cref="Drawing.DefaultRunProperties"/>. Shared by regular cChart
    /// and cx extended chart builders. Unspecified fields keep their
    /// defaults (size defaults to 10pt = 1000 hundredths).
    ///
    /// CONSISTENCY(chart-text-style): this is the single source of truth
    /// for parsing compound font specs. Callers wrap the result in their
    /// container of choice — <see cref="C.TextProperties"/> for regular
    /// cChart, or <c>CX.TxPrTextBody</c> for extended cxChart.
    /// </summary>
    internal static Drawing.DefaultRunProperties BuildDefaultRunPropertiesFromCompoundSpec(string spec)
    {
        var parts = spec.Split(':');
        // CONSISTENCY(pt-suffix): accept the unit-qualified form (`18pt`,
        // `10.5pt`) on input — without this, `int.TryParse("18pt")` failed
        // and silently defaulted to 1000 (10pt), so `axisFont=18pt:…` ignored
        // the size. Mirrors the project conventions "Font size input is lenient:
        // accepts `14`, `14pt`, `10.5pt`" rule.
        var sizeStr = parts.Length > 0
            ? (parts[0].EndsWith("pt", System.StringComparison.OrdinalIgnoreCase) ? parts[0][..^2] : parts[0])
            : "";
        var fontSize = sizeStr.Length > 0 && double.TryParse(sizeStr,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var fs)
            ? (int)System.Math.Round(fs * 100) : 1000;
        var color = parts.Length > 1 ? parts[1] : null;
        var fontName = parts.Length > 2 ? parts[2] : null;
        // CONSISTENCY(font-dash-form): when the spec has no colons and parts[0]
        // didn't parse as a number, try the dash form "name-size" (e.g.
        // "Arial-12", "Verdana-14") — what users naturally type when copying
        // a CSS-style font shorthand. Previously this silently dropped both
        // the typeface and the size (fontSize defaulted to 10pt, fontName=null,
        // no <a:latin> written). If there's no parseable size suffix we still
        // accept the bare typeface so `axisFont=Arial` writes the font.
        if (parts.Length == 1 && fontName == null)
        {
            var raw = parts[0];
            var dashIdx = raw.LastIndexOf('-');
            if (dashIdx > 0 && dashIdx < raw.Length - 1)
            {
                var maybeName = raw[..dashIdx];
                var maybeSize = raw[(dashIdx + 1)..];
                if (maybeSize.EndsWith("pt", System.StringComparison.OrdinalIgnoreCase))
                    maybeSize = maybeSize[..^2];
                if (double.TryParse(maybeSize, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var fs2))
                {
                    fontName = maybeName;
                    fontSize = (int)System.Math.Round(fs2 * 100);
                }
                else
                {
                    // Bare typeface ("Arial") — keep default size, set typeface.
                    fontName = raw;
                }
            }
            else if (sizeStr.Length > 0 && !double.TryParse(sizeStr,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out _))
            {
                // No dash, parts[0] isn't numeric → treat as bare typeface.
                fontName = raw;
            }
        }

        var defRp = new Drawing.DefaultRunProperties { FontSize = fontSize };
        if (!string.IsNullOrEmpty(color))
        {
            var solidFill = new Drawing.SolidFill();
            solidFill.AppendChild(BuildChartColorElement(color));
            defRp.AppendChild(solidFill);
        }
        if (!string.IsNullOrEmpty(fontName))
        {
            defRp.AppendChild(new Drawing.LatinFont { Typeface = fontName });
            defRp.AppendChild(new Drawing.EastAsianFont { Typeface = fontName });
        }
        return defRp;
    }

    /// <summary>
    /// Apply run-level styling from `{prefix}.color`/`{prefix}.size`/
    /// `{prefix}.font`/`{prefix}.bold` properties (and dotless aliases
    /// `{prefix}color`, `{prefix}size`, ...) onto an existing
    /// <see cref="Drawing.RunProperties"/>. Shared by both chart families.
    ///
    /// CONSISTENCY(chart-text-style): same vocabulary as
    /// <c>ChartHelper.Setter.cs</c> case `"title.color"`. Setter keeps its
    /// own inline implementation because it layers extra effects (glow /
    /// shadow) that are out of scope here.
    /// </summary>
    internal static void ApplyRunStyleProperties(Drawing.RunProperties rPr,
        Dictionary<string, string> properties, string keyPrefix)
    {
        string? Get(string suffix)
        {
            if (properties.TryGetValue($"{keyPrefix}.{suffix}", out var v) && !string.IsNullOrEmpty(v)) return v;
            if (properties.TryGetValue($"{keyPrefix}{suffix}", out v) && !string.IsNullOrEmpty(v)) return v;
            return null;
        }

        var color = Get("color");
        if (!string.IsNullOrEmpty(color))
        {
            rPr.RemoveAllChildren<Drawing.SolidFill>();
            var sf = new Drawing.SolidFill();
            sf.AppendChild(BuildChartColorElement(color));
            rPr.AppendChild(sf);
        }

        var size = Get("size");
        if (!string.IsNullOrEmpty(size))
        {
            var sizeStr = size.EndsWith("pt", StringComparison.OrdinalIgnoreCase) ? size[..^2] : size;
            if (double.TryParse(sizeStr, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var pts))
                rPr.FontSize = (int)Math.Round(pts * 100);
        }

        var font = Get("font");
        if (!string.IsNullOrEmpty(font))
        {
            rPr.RemoveAllChildren<Drawing.LatinFont>();
            rPr.RemoveAllChildren<Drawing.EastAsianFont>();
            rPr.AppendChild(new Drawing.LatinFont { Typeface = font });
            rPr.AppendChild(new Drawing.EastAsianFont { Typeface = font });
        }

        var bold = Get("bold");
        if (!string.IsNullOrEmpty(bold))
            rPr.Bold = ParseHelpers.IsTruthy(bold);
    }

    /// <summary>
    /// Apply text properties (font, size, color) to all axis labels.
    /// Format: "size:color:fontname" e.g. "10:8B949E:Helvetica Neue" or "10:CCCCCC".
    /// Used by the regular cChart path; delegates parsing to
    /// <see cref="BuildDefaultRunPropertiesFromCompoundSpec"/>.
    /// </summary>
    internal static void ApplyAxisTextProperties(OpenXmlCompositeElement axis, string value)
    {
        axis.RemoveAllChildren<C.TextProperties>();
        var defRp = BuildDefaultRunPropertiesFromCompoundSpec(value);

        var tp = new C.TextProperties(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph(new Drawing.ParagraphProperties(defRp))
        );

        // Insert before C.CrossingAxis or at end
        var crossAxis = axis.GetFirstChild<C.CrossingAxis>();
        if (crossAxis != null)
            axis.InsertBefore(tp, crossAxis);
        else
            axis.AppendChild(tp);
    }

    /// <summary>
    /// R15-4: set tick-label rotation on a category/value/date axis. Reuses
    /// the existing c:txPr subtree if any (preserves axisfont) and sets
    /// a:bodyPr/@rot. Creates a minimal c:txPr otherwise.
    /// </summary>
    internal static void ApplyAxisLabelRotation(OpenXmlCompositeElement axis, string rotAttrVal)
    {
        var tp = axis.GetFirstChild<C.TextProperties>();
        if (tp == null)
        {
            tp = new C.TextProperties(
                new Drawing.BodyProperties { Rotation = int.Parse(rotAttrVal) },
                new Drawing.ListStyle(),
                // CT_TextParagraph: pPr?, (br|r|fld)*, endParaRPr? — endParaRPr
                // is a sibling of pPr, NOT a child. Nesting it inside pPr
                // produces a schema-invalid file (pPr does not allow
                // endParaRPr as a child).
                new Drawing.Paragraph(
                    new Drawing.ParagraphProperties(),
                    new Drawing.EndParagraphRunProperties { Language = "en-US" })
            );
            var crossAxis = axis.GetFirstChild<C.CrossingAxis>();
            if (crossAxis != null)
                axis.InsertBefore(tp, crossAxis);
            else
                axis.AppendChild(tp);
            return;
        }
        var bodyPr = tp.GetFirstChild<Drawing.BodyProperties>();
        if (bodyPr == null)
        {
            bodyPr = new Drawing.BodyProperties { Rotation = int.Parse(rotAttrVal) };
            tp.PrependChild(bodyPr);
        }
        else
        {
            bodyPr.Rotation = int.Parse(rotAttrVal);
        }
    }

    /// <summary>
    /// Build a color element supporting both hex RGB and scheme color names.
    /// </summary>
    private static OpenXmlElement BuildChartColorElement(string value)
    {
        var schemeColor = value.ToLowerInvariant().TrimStart('#') switch
        {
            "accent1" => Drawing.SchemeColorValues.Accent1,
            "accent2" => Drawing.SchemeColorValues.Accent2,
            "accent3" => Drawing.SchemeColorValues.Accent3,
            "accent4" => Drawing.SchemeColorValues.Accent4,
            "accent5" => Drawing.SchemeColorValues.Accent5,
            "accent6" => Drawing.SchemeColorValues.Accent6,
            "dk1" or "dark1" => Drawing.SchemeColorValues.Dark1,
            "dk2" or "dark2" => Drawing.SchemeColorValues.Dark2,
            "lt1" or "light1" => Drawing.SchemeColorValues.Light1,
            "lt2" or "light2" => Drawing.SchemeColorValues.Light2,
            // The remaining ST_SchemeColorVal members. The reader emits these
            // verbatim (e.g. <a:schemeClr val="bg1"> → chartFill=bg1), so the
            // setter must accept them or the chart fill fails to round-trip.
            "bg1" or "background1" => Drawing.SchemeColorValues.Background1,
            "bg2" or "background2" => Drawing.SchemeColorValues.Background2,
            "tx1" or "text1" => Drawing.SchemeColorValues.Text1,
            "tx2" or "text2" => Drawing.SchemeColorValues.Text2,
            "phclr" => Drawing.SchemeColorValues.PhColor,
            "hlink" or "hyperlink" => Drawing.SchemeColorValues.Hyperlink,
            "folhlink" or "followedhyperlink" => Drawing.SchemeColorValues.FollowedHyperlink,
            _ => (Drawing.SchemeColorValues?)null
        };
        if (schemeColor.HasValue)
            return new Drawing.SchemeColor { Val = schemeColor.Value };
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml(value);
        var el = new Drawing.RgbColorModelHex { Val = rgb };
        if (alpha.HasValue) el.AppendChild(new Drawing.Alpha { Val = alpha.Value });
        return el;
    }

    // ==================== Series Builders ====================

    internal static C.BarChartSeries BuildBarSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.BarChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            BuildSeriesText(name)
        );
        if (color != null) ApplySeriesColor(series, color);
        // Real PowerPoint renders a MISSING invertIfNegative as true —
        // negative bars come out white with an outline. Write the explicit
        // spec default so negatives keep the series fill.
        series.AppendChild(new C.InvertIfNegative { Val = false });
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.LineChartSeries BuildLineSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.LineChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            BuildSeriesText(name)
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.AreaChartSeries BuildAreaSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.AreaChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            BuildSeriesText(name)
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.PieChartSeries BuildPieSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.PieChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            BuildSeriesText(name)
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.ScatterChartSeries BuildScatterSeries(uint idx, string name,
        double[]? xValues, double[] yValues)
    {
        var series = new C.ScatterChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            BuildSeriesText(name)
        );

        if (xValues != null)
        {
            var xLit = new C.NumberLiteral(new C.PointCount { Val = (uint)xValues.Length });
            for (int i = 0; i < xValues.Length; i++)
                xLit.AppendChild(new C.NumericPoint(new C.NumericValue(xValues[i].ToString("G"))) { Index = (uint)i });
            series.AppendChild(new C.XValues(xLit));
        }

        var yLit = new C.NumberLiteral(new C.PointCount { Val = (uint)yValues.Length });
        for (int i = 0; i < yValues.Length; i++)
            yLit.AppendChild(new C.NumericPoint(new C.NumericValue(yValues[i].ToString("G"))) { Index = (uint)i });
        series.AppendChild(new C.YValues(yLit));

        return series;
    }

    // ==================== Data Builders ====================

    internal static C.CategoryAxisData BuildCategoryData(string[] categories)
    {
        var strLit = new C.StringLiteral(new C.PointCount { Val = (uint)categories.Length });
        for (int i = 0; i < categories.Length; i++)
            strLit.AppendChild(new C.StringPoint(new C.NumericValue(categories[i])) { Index = (uint)i });
        return new C.CategoryAxisData(strLit);
    }

    internal static C.Values BuildValues(double[] values)
    {
        var numLit = new C.NumberLiteral(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)values.Length }
        );
        for (int i = 0; i < values.Length; i++)
            numLit.AppendChild(new C.NumericPoint(new C.NumericValue(values[i].ToString("G"))) { Index = (uint)i });
        return new C.Values(numLit);
    }

    /// <summary>
    /// Rewrite the SeriesText (c:tx) on a series so its content is a
    /// <c:strRef><c:f>formula</c:f>[<c:strCache>...]</c:strRef> referencing a
    /// single cell, instead of a literal <c:v>string</c:v>. Used when users pass
    /// series{N}.name=Sheet1!A1 — the legend/tooltip should resolve to the cell's
    /// current value, not show "Sheet1!A1" as literal text.
    ///
    /// If cachedValue is non-null, a minimal c:strCache with one c:pt idx="0" is
    /// attached so first-open viewers (before Excel recalculates) still see the
    /// resolved text. When null, Excel fills the cache on open.
    /// </summary>
    internal static void RewriteSeriesTextAsRef(
        OpenXmlCompositeElement series, string formula, string? cachedValue)
    {
        var serText = series.GetFirstChild<C.SeriesText>();
        if (serText == null) return;
        serText.RemoveAllChildren();
        var strRef = new C.StringReference(new C.Formula(formula));
        if (cachedValue != null)
        {
            var cache = new C.StringCache(
                new C.PointCount { Val = 1U },
                new C.StringPoint(new C.NumericValue(cachedValue)) { Index = 0U });
            strRef.AppendChild(cache);
        }
        serText.AppendChild(strRef);
    }

    /// <summary>
    /// Build a Values element with a NumberReference (cell range formula, no cache).
    /// </summary>
    internal static C.Values BuildValuesRef(string formula)
    {
        var numRef = new C.NumberReference(new C.Formula(formula));
        return new C.Values(numRef);
    }

    /// <summary>
    /// Build a CategoryAxisData element with a StringReference (cell range formula, no cache).
    /// </summary>
    internal static C.CategoryAxisData BuildCategoryDataRef(string formula)
    {
        var strRef = new C.StringReference(new C.Formula(formula));
        return new C.CategoryAxisData(strRef);
    }

    /// <summary>
    /// Convert a NumberLiteral to a NumberingCache so chart viewers can display
    /// cached values without recalculating cell references.
    /// </summary>
    private static C.NumberingCache? BuildNumberingCacheFromLiteral(
        C.NumberLiteral? literal, HashSet<int>? skipIndexes = null)
    {
        if (literal == null) return null;
        var points = literal.Elements<C.NumericPoint>().ToList();
        if (points.Count == 0) return null;
        var cache = new C.NumberingCache();
        var fmtCode = literal.GetFirstChild<C.FormatCode>();
        cache.AppendChild(new C.FormatCode(fmtCode?.Text ?? "General"));
        var ptCount = literal.GetFirstChild<C.PointCount>();
        if (ptCount != null)
            cache.AppendChild(new C.PointCount { Val = ptCount.Val });
        foreach (var pt in points)
        {
            // R20-03: under dispBlanksAs=gap, omit points at blank source
            // indexes so Excel renders a gap (line break) instead of 0.
            if (skipIndexes != null && pt.Index?.Value is uint idx && skipIndexes.Contains((int)idx))
                continue;
            cache.AppendChild((C.NumericPoint)pt.CloneNode(true));
        }
        return cache;
    }

    /// <summary>
    /// Convert a StringLiteral to a StringCache so chart viewers can display
    /// cached labels without recalculating cell references.
    /// </summary>
    private static C.StringCache? BuildStringCacheFromLiteral(C.StringLiteral? literal)
    {
        if (literal == null) return null;
        var points = literal.Elements<C.StringPoint>().ToList();
        if (points.Count == 0) return null;
        var cache = new C.StringCache();
        var ptCount = literal.GetFirstChild<C.PointCount>();
        if (ptCount != null)
            cache.AppendChild(new C.PointCount { Val = ptCount.Val });
        foreach (var pt in points)
            cache.AppendChild((C.StringPoint)pt.CloneNode(true));
        return cache;
    }

    // ==================== Axis Builders ====================

    internal static C.CategoryAxis BuildCategoryAxis(uint axisId, uint crossAxisId)
    {
        return new C.CategoryAxis(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.MajorTickMark { Val = C.TickMarkValues.Outside },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = crossAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true },
            new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset { Val = 100 }
        );
    }

    internal static C.ValueAxis BuildValueAxis(uint axisId, uint crossAxisId, C.AxisPositionValues position)
    {
        return new C.ValueAxis(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = position },
            new C.MajorGridlines(),
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
            new C.MajorTickMark { Val = C.TickMarkValues.Outside },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = crossAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between }
        );
    }

    /// <summary>
    /// Build a c:serAx (series axis) for 3D chart types. CT_Bar3DChart /
    /// CT_Line3DChart / CT_Area3DChart require a third axId on the chart
    /// element bound to a serAx in the plotArea — without it the chart fails
    /// OpenXmlValidator and PowerPoint repairs the file with a degenerate
    /// render (no series progression across the depth axis).
    /// </summary>
    internal static C.SeriesAxis BuildSeriesAxis(uint axisId, uint crossAxisId)
    {
        return new C.SeriesAxis(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.CrossingAxis { Val = crossAxisId }
        );
    }

    // Shared helper for chart series text: validate XML-illegal chars up-front
    // so callers don't have to remember (every `BuildSeriesText(name)`
    // call site was previously a leak point — R32). Returns the SeriesText
    // wrapping a NumericValue, ready to AppendChild.
    internal static C.SeriesText BuildSeriesText(string name)
    {
        OfficeCli.Core.ParseHelpers.ValidateXmlText(name, "series name");
        return new C.SeriesText(new C.NumericValue(name));
    }

    // ==================== Title Builder ====================

    internal static C.Title BuildChartTitle(string titleText, string? titleLang = null)
    {
        // CONSISTENCY(chart-cell-ref): if titleText looks like a single-cell
        // reference (e.g. "Sheet1!A1"), emit <c:tx><c:strRef> so Excel resolves
        // the cell on open. Same fix family as R17-B1 (series name strRef).
        // Applies to chart title and cat/val axis titles (R18-B1/B2).
        // Accept the natural Excel spelling with a leading '=' (=Sheet1!A1);
        // without the strip it fell through to the literal-text branch and
        // the chart displayed the formula string verbatim.
        var titleRef = titleText.TrimStart();
        if (titleRef.StartsWith('=')) titleRef = titleRef[1..];
        if (IsCellReference(titleRef))
        {
            var formula = NormalizeCellReference(titleRef);
            return new C.Title(
                new C.ChartText(
                    new C.StringReference(new C.Formula(formula))
                ),
                new C.Overlay { Val = false }
            );
        }

        // R53 tester-2: source charts authored in zh-CN / ja-JP / ko-KR carry
        // <a:rPr lang="zh-CN"/> on the title run; dump→replay hard-coded
        // "en-US" and the locale-specific text shaping (line break rules,
        // font fallback) regressed. Accept the source lang via title.lang
        // input and default to en-US for new charts.
        var titleLangVal = string.IsNullOrEmpty(titleLang) ? "en-US" : titleLang!;
        return new C.Title(
            new C.ChartText(
                new C.RichText(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties(
                            new Drawing.DefaultRunProperties { FontSize = 1400, Bold = true }
                        ),
                        new Drawing.Run(
                            new Drawing.RunProperties { Language = titleLangVal, FontSize = 1400, Bold = true },
                            new Drawing.Text(titleText)
                        )
                    )
                )
            ),
            new C.Overlay { Val = false }
        );
    }

}
