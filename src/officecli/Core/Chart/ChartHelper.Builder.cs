// Copyright 2025 OfficeCli (officecli.ai)
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
        var chart = new C.Chart();

        if (!string.IsNullOrEmpty(title))
            chart.AppendChild(BuildChartTitle(title));

        var originalCategories = categories;
        if (categories == null && seriesData.Count > 0)
        {
            var maxLen = seriesData.Max(s => s.values.Length);
            categories = Enumerable.Range(1, maxLen).Select(i => i.ToString()).ToArray();
        }

        var plotArea = new C.PlotArea(new C.Layout());
        uint catAxisId = 1;
        uint valAxisId = 2;

        OpenXmlCompositeElement? chartElement;
        bool needsAxes = true;

        var colors = ParseSeriesColors(properties);

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
                chartElement = bar3dAuto;
                break;
            }
            case "bar":
                chartElement = BuildBarChart(C.BarDirectionValues.Bar, stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
                break;
            case "column":
                chartElement = BuildBarChart(C.BarDirectionValues.Column, stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
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
                    var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
                    line3d.AppendChild(BuildLineSeries((uint)i, seriesData[i].name,
                        categories, seriesData[i].values, color));
                }
                line3d.AppendChild(new C.AxisId { Val = catAxisId });
                line3d.AppendChild(new C.AxisId { Val = valAxisId });
                chartElement = line3d;
                break;
            }
            case "line":
                chartElement = BuildLineChart(stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
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
                    var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
                    area3d.AppendChild(BuildAreaSeries((uint)i, seriesData[i].name,
                        categories, seriesData[i].values, color));
                }
                area3d.AppendChild(new C.AxisId { Val = catAxisId });
                area3d.AppendChild(new C.AxisId { Val = valAxisId });
                chartElement = area3d;
                break;
            }
            case "area":
                chartElement = BuildAreaChart(stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
                break;
            case "pie" when is3D:
            {
                var pie3d = new C.Pie3DChart(
                    new C.VaryColors { Val = true }
                );
                for (int i = 0; i < seriesData.Count; i++)
                {
                    var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
                    pie3d.AppendChild(BuildPieSeries((uint)i, seriesData[i].name,
                        categories, seriesData[i].values, color));
                }
                chartElement = pie3d;
                needsAxes = false;
                break;
            }
            case "pie":
                chartElement = BuildPieChart(categories, seriesData, colors);
                needsAxes = false;
                break;
            case "doughnut":
                chartElement = BuildDoughnutChart(categories, seriesData, colors);
                needsAxes = false;
                break;
            case "scatter":
                var scatterStyle = properties.GetValueOrDefault("scatterStyle", "lineMarker");
                chartElement = BuildScatterChart(categories, seriesData, catAxisId, valAxisId, scatterStyle);
                break;
            case "bubble":
                chartElement = BuildBubbleChart(categories, seriesData, catAxisId, valAxisId, colors);
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
                return wfChartSpace;
            }
            case "combo":
            {
                int splitAt = 1;
                if (properties.TryGetValue("combosplit", out var splitStr))
                    splitAt = ParseHelpers.SafeParseInt(splitStr, "combosplit");
                splitAt = Math.Min(splitAt, seriesData.Count);

                var barData = seriesData.Take(splitAt).ToList();
                var lineData = seriesData.Skip(splitAt).ToList();

                var comboBar = new C.BarChart(
                    new C.BarDirection { Val = C.BarDirectionValues.Column },
                    new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                    new C.VaryColors { Val = false }
                );
                for (int ci = 0; ci < barData.Count; ci++)
                {
                    var clr = colors != null && ci < colors.Length ? colors[ci] : DefaultSeriesColors[ci % DefaultSeriesColors.Length];
                    comboBar.AppendChild(BuildBarSeries((uint)ci, barData[ci].name, categories, barData[ci].values, clr));
                }
                comboBar.AppendChild(new C.AxisId { Val = catAxisId });
                comboBar.AppendChild(new C.AxisId { Val = valAxisId });
                plotArea.AppendChild(comboBar);

                if (lineData.Count > 0)
                {
                    var comboLine = new C.LineChart(
                        new C.Grouping { Val = C.GroupingValues.Standard },
                        new C.VaryColors { Val = false }
                    );
                    for (int ci = 0; ci < lineData.Count; ci++)
                    {
                        var sIdx = (uint)(splitAt + ci);
                        var cIdx = splitAt + ci;
                        var clr = colors != null && cIdx < colors.Length ? colors[cIdx] : DefaultSeriesColors[cIdx % DefaultSeriesColors.Length];
                        comboLine.AppendChild(BuildLineSeries(sIdx, lineData[ci].name, categories, lineData[ci].values, clr));
                    }
                    comboLine.AppendChild(new C.ShowMarker { Val = true });
                    comboLine.AppendChild(new C.AxisId { Val = catAxisId });
                    comboLine.AppendChild(new C.AxisId { Val = valAxisId });
                    plotArea.AppendChild(comboLine);
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
            }
        }

        chart.AppendChild(plotArea);

        var showLegend = properties.GetValueOrDefault("legend", "true");
        // CONSISTENCY(legend-hide-alias / R34-1): accept hide=true / hidden=true
        // as aliases for legend=none so users with a "hide it" mental model
        // don't reach for legend=hidden (which is now rejected).
        if ((properties.TryGetValue("hide", out var hideVal) && ParseHelpers.IsTruthy(hideVal)) ||
            (properties.TryGetValue("hidden", out var hiddenVal) && ParseHelpers.IsTruthy(hiddenVal)))
        {
            showLegend = "none";
        }
        // Bare "true" keeps the documented default of bottom.
        if (showLegend.Equals("true", StringComparison.OrdinalIgnoreCase))
            showLegend = "bottom";
        if (!showLegend.Equals("false", StringComparison.OrdinalIgnoreCase) &&
            !showLegend.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            // CONSISTENCY(strict-enums / R34-1): shared helper rejects
            // unknown positions with the documented valid set.
            var legendPos = ParseLegendPosition(showLegend);
            chart.AppendChild(new C.Legend(
                new C.LegendPosition { Val = legendPos },
                new C.Overlay { Val = false }
            ));
        }

        chart.AppendChild(new C.PlotVisibleOnly { Val = true });
        chart.AppendChild(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });

        chartSpace.AppendChild(chart);

        // Apply cell references for dotted syntax (series1.values=Sheet1!B2:B13)
        ApplySeriesReferences(plotArea, properties);

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
        if (extSeries != null && !extSeries.Any(s => s.ValuesRef != null || s.CategoriesRef != null) && !hasNameRef)
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
                        var strCache = BuildStringCacheFromLiteral(catEl.GetFirstChild<C.StringLiteral>());
                        catEl.RemoveAllChildren();
                        var strRef = new C.StringReference(new C.Formula(catRef));
                        if (strCache != null)
                            strRef.AppendChild(strCache);
                        catEl.AppendChild(strRef);
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
        "plotareafill", "plotfill", "chartareafill", "chartfill",
        "linewidth", "linedash", "dash", "marker", "markers", "markersize",
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
        "explosion", "explode", "invertifneg", "invertifnegative",
        "errbars", "errorbars", "series.shadow", "seriesshadow",
        "series.outline", "seriesoutline",
        "bubblescale", "shownegbubbles", "sizerepresents",
        "gapdepth", "shape", "barshape",
        "droplines", "hilowlines", "updownbars", "serlines", "serieslines",
        "axisorientation", "axisreverse", "logbase", "logscale", "yaxisscale",
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
        // Area fill
        "areafill", "area.fill",
    };

    /// <summary>
    /// Prefixes for dynamic deferred keys (e.g. title.x, plotArea.y, legend.w,
    /// dataLabel1.text, dataTable.show*, displayUnitsLabel.*, trendlineLabel.*).
    /// </summary>
    private static readonly string[] DeferredPrefixes =
    [
        "title.", "plotarea.", "legend.", "datalabel",
        "datatable.", "displayunitslabel.", "trendlinelabel."
    ];

    /// <summary>
    /// Check if a property key should be deferred from BuildChartSpace to SetChartProperties.
    /// Matches exact keys in <see cref="DeferredAddKeys"/> plus dynamic prefix patterns.
    /// </summary>
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
        if (TryParseSeriesDottedKey(key, out _, out var sProp)
            && DeferredSeriesSubkeys.Contains(sProp)) return true;
        return false;
    }

    // Per-series dotted subkeys that route through HandleSeriesDottedProperty
    // during SetChartProperties (post-build). See IsDeferredKey.
    private static readonly HashSet<string> DeferredSeriesSubkeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "gradient", "gradientfill",
        "smooth", "trendline", "marker", "markersize",
        "invertifneg", "invertifnegative",
        "errbars", "errorbars",
        "explosion", "explode",
        "linewidth", "linedash", "dash",
        "shadow", "outline",
        "alpha", "transparency",
    };

    // ==================== Chart Type Builders ====================

    internal static C.BarChart BuildBarChart(
        C.BarDirectionValues direction, bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null)
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
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
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
        uint catAxisId, uint valAxisId, string[]? colors = null)
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
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            lineChart.AppendChild(BuildLineSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values, color));
        }

        lineChart.AppendChild(new C.ShowMarker { Val = true });
        lineChart.AppendChild(new C.AxisId { Val = catAxisId });
        lineChart.AppendChild(new C.AxisId { Val = valAxisId });
        return lineChart;
    }

    internal static C.AreaChart BuildAreaChart(
        bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null)
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
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
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

    internal static C.DoughnutChart BuildDoughnutChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        string[]? colors = null)
    {
        var chart = new C.DoughnutChart(new C.VaryColors { Val = true });
        if (seriesData.Count > 0)
        {
            var series = BuildPieSeries(0, seriesData[0].name,
                categories, seriesData[0].values);
            ApplyDataPointColors(series, seriesData[0].values.Length, colors);
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
        uint catAxisId, uint valAxisId, string scatterStyle = "lineMarker")
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
            // For marker-only style, explicitly hide connecting lines
            if (hideLines)
            {
                var spPr = ser.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
                if (ser.GetFirstChild<C.ChartShapeProperties>() == null) ser.AppendChild(spPr);
                spPr.RemoveAllChildren<Drawing.Outline>();
                spPr.AppendChild(new Drawing.Outline(new Drawing.NoFill()));
            }
            scatterChart.AppendChild(ser);
        }

        scatterChart.AppendChild(new C.AxisId { Val = catAxisId });
        scatterChart.AppendChild(new C.AxisId { Val = valAxisId });
        return scatterChart;
    }

    // ==================== Bubble Chart ====================

    internal static C.BubbleChart BuildBubbleChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null)
    {
        var bubbleChart = new C.BubbleChart(new C.VaryColors { Val = false });

        double[]? xValues = null;
        if (categories != null)
            xValues = categories.Select(c => double.TryParse(c, out var v) ? v : 0).ToArray();

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            var (name, values) = seriesData[i];
            var series = new C.BubbleChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                new C.SeriesText(new C.NumericValue(name))
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

            // Bubble sizes — use the values as sizes by default, or a third series if provided
            var sizeLit = new C.NumberLiteral(new C.PointCount { Val = (uint)values.Length });
            for (int j = 0; j < values.Length; j++)
                sizeLit.AppendChild(new C.NumericPoint(new C.NumericValue(values[j].ToString("G"))) { Index = (uint)j });
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
                new C.SeriesText(new C.NumericValue(seriesData[i].name))
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
        var stockChart = new C.StockChart();

        for (int i = 0; i < seriesData.Count; i++)
        {
            var series = new C.LineChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                new C.SeriesText(new C.NumericValue(seriesData[i].name))
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

    // ==================== Series Color ====================

    internal static void ApplySeriesColor(OpenXmlCompositeElement series, string color)
    {
        series.RemoveAllChildren<C.ChartShapeProperties>();
        var spPr = new C.ChartShapeProperties();
        var solidFill = new Drawing.SolidFill();
        solidFill.AppendChild(BuildChartColorElement(color));
        spPr.AppendChild(solidFill);

        // For line/scatter series, also set a:ln so Excel uses the correct stroke color
        var parentName = series.Parent?.LocalName;
        if (parentName is "lineChart" or "scatterChart" or "radarChart")
        {
            const int defaultStrokeWidthEmu = 25400; // 2pt × 12700 EMU/pt
            var outline = new Drawing.Outline { Width = defaultStrokeWidthEmu };
            var lnFill = new Drawing.SolidFill();
            lnFill.AppendChild(BuildChartColorElement(color));
            outline.AppendChild(lnFill);
            spPr.AppendChild(outline);
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
    /// </summary>
    private static OpenXmlElement BuildFillElement(string value)
    {
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
            return new Drawing.NoFill();

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
            gradFill.AppendChild(new Drawing.LinearGradientFill { Angle = anglePart * 60000, Scaled = true });
            return gradFill;
        }

        // Solid fill
        var solidFill = new Drawing.SolidFill();
        solidFill.AppendChild(BuildChartColorElement(value));
        return solidFill;
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
        var fontSize = parts.Length > 0 && int.TryParse(parts[0], out var fs) ? fs * 100 : 1000;
        var color = parts.Length > 1 ? parts[1] : null;
        var fontName = parts.Length > 2 ? parts[2] : null;

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
                new Drawing.Paragraph(new Drawing.ParagraphProperties(new Drawing.EndParagraphRunProperties { Language = "en-US" }))
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
            new C.SeriesText(new C.NumericValue(name))
        );
        if (color != null) ApplySeriesColor(series, color);
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
            new C.SeriesText(new C.NumericValue(name))
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
            new C.SeriesText(new C.NumericValue(name))
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
            new C.SeriesText(new C.NumericValue(name))
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
            new C.SeriesText(new C.NumericValue(name))
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

    // ==================== Title Builder ====================

    internal static C.Title BuildChartTitle(string titleText)
    {
        // CONSISTENCY(chart-cell-ref): if titleText looks like a single-cell
        // reference (e.g. "Sheet1!A1"), emit <c:tx><c:strRef> so Excel resolves
        // the cell on open. Same fix family as R17-B1 (series name strRef).
        // Applies to chart title and cat/val axis titles (R18-B1/B2).
        if (IsCellReference(titleText))
        {
            var formula = NormalizeCellReference(titleText);
            return new C.Title(
                new C.ChartText(
                    new C.StringReference(new C.Formula(formula))
                ),
                new C.Overlay { Val = false }
            );
        }

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
                            new Drawing.RunProperties { Language = "en-US", FontSize = 1400, Bold = true },
                            new Drawing.Text(titleText)
                        )
                    )
                )
            ),
            new C.Overlay { Val = false }
        );
    }

}
