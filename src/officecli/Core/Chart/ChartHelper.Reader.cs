// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

internal static partial class ChartHelper
{
    // ==================== Chart Readback ====================

    // BUG-DUMP-CHART-AXID-UNSIGNED: c:axId/@val is xsd:unsignedInt, but Word
    // routinely emits values >= 2^31 in their signed-overflow text form (e.g.
    // "-1880390128" = unsigned 2414577168). UInt32Value.Value throws a
    // FormatException on the negative string, which previously crashed the ENTIRE
    // document dump (Error: input string '-1880390128' was not in a correct
    // format). Read the raw text and reinterpret leniently so axis-rank mapping
    // (the only consumer) survives; Word opens such files fine.
    private static uint? SafeAxisIdVal(C.AxisId? ax)
    {
        var raw = ax?.Val?.InnerText;
        if (string.IsNullOrEmpty(raw)) return null;
        if (uint.TryParse(raw, out var u)) return u;
        if (long.TryParse(raw, out var l)) return unchecked((uint)l);
        return null;
    }

    internal static void ReadChartProperties(C.Chart chart, DocumentNode node, int depth)
    {
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return;

        // R16-bt-2 — chart reading direction. Setter stamps rtl on
        // chartSpace c:txPr/a:lstStyle/a:lvl1pPr (and propagates to
        // axis/legend/dLbls). Surface the chart-level value as the
        // canonical "direction" key, mirroring shape/textbox readback.
        if (chart.Parent is C.ChartSpace chartSpace)
        {
            var rootTxPr = chartSpace.GetFirstChild<C.TextProperties>();
            var rootLvl1 = rootTxPr?.GetFirstChild<Drawing.ListStyle>()
                ?.GetFirstChild<Drawing.Level1ParagraphProperties>();
            if (rootLvl1?.RightToLeft?.HasValue == true)
                node.Format["direction"] = rootLvl1.RightToLeft.Value ? "rtl" : "ltr";

            // chartSpace-level default text properties (<c:txPr> directly
            // under c:chartSpace) set the base font for EVERY chart element
            // — an 18pt default reshapes axis labels, legend and plot-area
            // proportions. Carry it verbatim (*Raw family) so dump→replay
            // preserves the base font.
            if (rootTxPr != null)
                node.InternalFormat["chartTxPrRaw"] = rootTxPr.OuterXml;

            // 3D wall/floor elements — sources hide the wall grid with
            // <a:ln><a:noFill/> spPr; without carrying them the rebuilt 3D
            // chart shows PowerPoint's default wall outlines.
            var chartForWalls = chartSpace.GetFirstChild<C.Chart>();
            if (chartForWalls != null)
            {
                var floorEl2 = chartForWalls.GetFirstChild<C.Floor>();
                if (floorEl2?.HasChildren == true)
                    node.InternalFormat["floorRaw"] = floorEl2.InnerXml;
                var sideWallEl = chartForWalls.GetFirstChild<C.SideWall>();
                if (sideWallEl?.HasChildren == true)
                    node.InternalFormat["sideWallRaw"] = sideWallEl.InnerXml;
                var backWallEl = chartForWalls.GetFirstChild<C.BackWall>();
                if (backWallEl?.HasChildren == true)
                    node.InternalFormat["backWallRaw"] = backWallEl.InnerXml;
            }

            // 1904 date epoch flag — Builder always writes an explicit
            // <c:date1904>, so only the non-default true needs surfacing.
            if (chartSpace.GetFirstChild<C.Date1904>()?.Val?.Value == true)
                node.Format["date1904"] = "true";

            // Chart style number (<c:style val="N"/>, or the mc:Fallback
            // form when the source wraps it in AlternateContent). Drives
            // gridline tint / effect defaults in real PowerPoint.
            var styleEl = chartSpace.Elements<C.Style>().FirstOrDefault()
                ?? chartSpace.Descendants<C.Style>().FirstOrDefault();
            if (styleEl?.Val?.HasValue == true)
                node.Format["chartStyle"] = styleEl.Val.Value.ToString();
        }

        var chartType = DetectChartType(plotArea);
        if (chartType != null) node.Format["chartType"] = chartType;

        // Waterfall: surface increase/decrease/totalColor at chart level so
        // dump→replay preserves the bar colors. Without these the encoded
        // triplet (Base/Increase/Decrease) is collapsed back to deltas by
        // the emitter and Builder falls back to the default 4472C4 / FF0000
        // palette, dropping the user's customisation.
        if (chartType == "waterfall"
            && plotArea.GetFirstChild<C.BarChart>() is C.BarChart wfBar)
        {
            var wfSeries = wfBar.Elements<C.BarChartSeries>().ToList();
            // Increase = series[1], Decrease = series[2] (Builder convention).
            if (wfSeries.Count >= 3)
            {
                var incFill = wfSeries[1].GetFirstChild<C.ChartShapeProperties>()
                    ?.GetFirstChild<Drawing.SolidFill>();
                var incClr = incFill != null ? ReadColorFromFill(incFill) : null;
                if (incClr != null) node.Format["increaseColor"] = incClr;

                var decFill = wfSeries[2].GetFirstChild<C.ChartShapeProperties>()
                    ?.GetFirstChild<Drawing.SolidFill>();
                var decClr = decFill != null ? ReadColorFromFill(decFill) : null;
                if (decClr != null) node.Format["decreaseColor"] = decClr;

                // Total bar = last DataPoint override on Increase series.
                var dpts = wfSeries[1].Elements<C.DataPoint>().ToList();
                var lastDpt = dpts.LastOrDefault();
                var totFill = lastDpt?.GetFirstChild<C.ChartShapeProperties>()
                    ?.GetFirstChild<Drawing.SolidFill>();
                var totClr = totFill != null ? ReadColorFromFill(totFill) : null;
                if (totClr != null) node.Format["totalColor"] = totClr;
            }
        }

        // R24 — for combo charts surface the per-series type list (and the
        // split point if it cleanly partitions into a primary block + tail)
        // so dump→replay can reconstruct mixed-type charts. Without this,
        // every combo collapsed back to a column+line split at index 1.
        if (chartType == "combo")
        {
            var typesPerSeries = new List<string>();
            foreach (var ct in plotArea.Elements<OpenXmlCompositeElement>())
            {
                string? ctLabel = ct switch
                {
                    C.BarChart bc => bc.GetFirstChild<C.BarDirection>()?.Val?.Value == C.BarDirectionValues.Bar
                        ? "bar" : "column",
                    C.LineChart => "line",
                    C.AreaChart => "area",
                    C.ScatterChart => "scatter",
                    C.PieChart => "pie",
                    C.DoughnutChart => "doughnut",
                    C.BubbleChart => "bubble",
                    C.RadarChart => "radar",
                    _ => null,
                };
                if (ctLabel == null) continue;
                // Grouping-qualified tokens (columnstacked / areapercentstacked
                // …) so a combo whose groups are stacked doesn't replay as
                // clustered/standard. BuildComboGroup parses the suffix back.
                if (ctLabel is "column" or "bar" or "area" or "line")
                {
                    var grp = ct is C.BarChart bch2
                        ? bch2.GetFirstChild<C.BarGrouping>()?.Val?.InnerText
                        : ct.GetFirstChild<C.Grouping>()?.Val?.InnerText;
                    if (grp == "stacked") ctLabel += "stacked";
                    else if (grp == "percentStacked") ctLabel += "percentstacked";
                }
                var serCount = ct.Elements<OpenXmlCompositeElement>()
                    .Count(e => e.LocalName == "ser");
                for (int i = 0; i < serCount; i++) typesPerSeries.Add(ctLabel);
            }
            if (typesPerSeries.Count > 0)
            {
                node.Format["comboTypes"] = string.Join(",", typesPerSeries);
                // combosplit = number of leading series of the first type — the
                // partition the simple Builder.combo path can rebuild without
                // touching RebuildComboChart.
                int splitAt = 0;
                var first = typesPerSeries[0];
                while (splitAt < typesPerSeries.Count && typesPerSeries[splitAt] == first)
                    splitAt++;
                if (splitAt > 0 && splitAt < typesPerSeries.Count)
                    node.Format["combosplit"] = splitAt;
            }
        }

        var titleEl = chart.GetFirstChild<C.Title>();
        // Concatenate ALL text runs — a styled title splits its text across
        // multiple <a:r> runs and taking only the first truncated it
        // ("Stacked column mixed with…" → "Stacked ").
        var titleRuns = titleEl?.Descendants<Drawing.Text>().Select(t => t.Text).ToList();
        var titleText = titleRuns is { Count: > 0 } ? string.Concat(titleRuns) : null;
        if (titleText == null && titleEl != null)
        {
            // BuildChartTitle routes single-cell-reference values (e.g. "Q1",
            // "Sheet1!A1") through a <c:strRef><c:f>...</c:f></c:strRef> path
            // instead of <a:t> literal text. Surface the formula so a get→set
            // round-trip preserves the reference and the schema-declared
            // 'title' get readback isn't silently empty.
            var strRefFormula = titleEl.Descendants<C.Formula>().FirstOrDefault()?.Text;
            if (!string.IsNullOrEmpty(strRefFormula)) titleText = strRefFormula;

            // Auto-title: an empty <c:title> with autoTitleDeleted=0 makes
            // real PowerPoint title a SINGLE-series chart with the series
            // name. Surface that resolved name so dump→replay keeps the
            // rendered title (the rebuilt chart writes it as literal text).
            if (titleText == null
                && chart.GetFirstChild<C.AutoTitleDeleted>()?.Val?.Value != true)
            {
                var serEls = plotArea.Descendants<OpenXmlCompositeElement>()
                    .Where(e => e.LocalName == "ser").ToList();
                if (serEls.Count == 1)
                {
                    var serName = serEls[0].GetFirstChild<C.SeriesText>()
                        ?.Descendants<C.NumericValue>().FirstOrDefault()?.Text;
                    if (!string.IsNullOrEmpty(serName)) titleText = serName;
                }
                // Multi-series auto-title: PowerPoint renders its localized
                // "Chart Title" placeholder. No literal text can reproduce
                // that locale-dependent string — signal the builder to write
                // an empty <c:title/> + autoTitleDeleted=0 instead.
                if (titleText == null)
                    node.Format["autoTitle"] = "true";
            }
        }
        if (titleText != null) node.Format["title"] = titleText;

        // Title overlay (<c:title><c:overlay val="1"/></c:title>) — when true,
        // the title is drawn on top of the plot area instead of reserving
        // space above it. BuildChartTitle defaults to overlay=false, so only
        // surface the truthy form (mirrors `legend.overlay` and
        // `autoTitleDeleted`) for dump→replay round-trip fidelity. Without
        // this emit, source charts authored with title-on-plot lost the
        // overlay flag silently via SDK default on replay.
        if (titleEl != null)
        {
            var titleOverlay = titleEl.GetFirstChild<C.Overlay>()?.Val;
            if (titleOverlay?.HasValue == true && titleOverlay.Value)
                node.Format["title.overlay"] = "true";
        }

        // AutoTitleDeleted only round-trips when explicitly emitted in the
        // OOXML — its absence is the default. Surface only the truthy form
        // so dump→replay doesn't fight scatter charts, which Excel writes
        // with <c:autoTitleDeleted val="1"/> to suppress the auto-generated
        // single-series title. Without this emit, replayed scatter charts
        // gained a synthetic title and PowerPoint flagged the file as
        // corrupt (Error 422).
        var autoTitleDeleted = chart.GetFirstChild<C.AutoTitleDeleted>()?.Val?.Value;
        if (autoTitleDeleted == true) node.Format["autoTitleDeleted"] = "true";

        // Reference lines (AddReferenceLine overlays) — emit as a single
        // chart-level `referenceLine=value:color:label:dash` (or semicolon-
        // joined list) so dump→replay reconstructs the same overlay.
        // Without this the lineChart sibling round-tripped as a real data
        // series and the chartType heuristic that excluded ref-line-only
        // LineCharts found nothing to emit, so the overlay was lost.
        {
            var refLines = ReadReferenceLines(plotArea);
            if (refLines.Count > 0)
            {
                var specs = refLines.Select(r =>
                {
                    var v = r.Value.ToString("G",
                        System.Globalization.CultureInfo.InvariantCulture);
                    var label = string.IsNullOrEmpty(r.Name) ? "" : r.Name;
                    var dash = r.Dash;
                    return $"{v}:{r.Color}:{label}:{dash}";
                });
                node.Format["referenceLine"] = string.Join(";", specs);
            }
        }

        // Title formatting: font, size, color, bold from RunProperties
        if (titleEl != null)
        {
            var titleRun = titleEl.Descendants<Drawing.Run>().FirstOrDefault();
            var titleRp = titleRun?.RunProperties;
            if (titleRp != null)
            {
                var titleFont = titleRp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
                if (titleFont != null) node.Format["title.font"] = titleFont;
                if (titleRp.FontSize?.HasValue == true)
                    node.Format["title.size"] = $"{titleRp.FontSize.Value / 100.0:0.##}pt";
                var titleFill = titleRp.GetFirstChild<Drawing.SolidFill>();
                if (titleFill != null)
                {
                    var tColor = ReadColorFromFill(titleFill);
                    if (tColor != null) node.Format["title.color"] = tColor;
                }
                if (titleRp.Bold?.HasValue == true)
                    node.Format["title.bold"] = titleRp.Bold.Value ? "true" : "false";
                // R53 tester-2: round-trip the title run's lang attribute
                // (zh-CN / ja-JP / ko-KR / en-US). Default-construction
                // hard-coded en-US, so a source-authored locale silently
                // regressed on dump→replay.
                if (titleRp.Language?.HasValue == true && !string.IsNullOrEmpty(titleRp.Language.Value))
                    node.Format["title.lang"] = titleRp.Language.Value!;
            }

            // BUG-DUMP-R35-1: the title font lives both on the run-level <a:rPr>
            // (which renders, captured above) AND on the paragraph-level
            // <a:pPr><a:defRPr> (algn + color + typeface). BuildChartTitle only
            // emits a bare defRPr (sz/bold), so the source's defRPr colour /
            // typeface / paragraph alignment were dropped on rebuild — a
            // render-neutral but real XML-fidelity gap. Capture the title's
            // <a:pPr> verbatim and inject it back, replacing the builder's
            // default. Only when the defRPr carries explicit styling (an
            // alignment attr alone is also worth round-tripping).
            var titlePPr = titleEl.Descendants<Drawing.ParagraphProperties>().FirstOrDefault();
            if (titlePPr != null)
            {
                var titleDefRp = titlePPr.GetFirstChild<Drawing.DefaultRunProperties>();
                bool defRpMeaningful = titleDefRp != null && (
                    titleDefRp.ChildElements.Any(c => c.LocalName is "solidFill"
                        or "latin" or "ea" or "cs")
                    || titleDefRp.GetAttributes().Any(a => a.LocalName is "i" or "u"));
                bool hasAlgn = titlePPr.GetAttributes().Any(a => a.LocalName == "algn");
                if (defRpMeaningful || hasAlgn)
                    node.InternalFormat["title.pPr"] = titlePPr.OuterXml;
            }
        }

        var legend = chart.GetFirstChild<C.Legend>();
        if (legend != null)
        {
            // Absent <c:legendPos> → ECMA-376 CT_LegendPos default is "r"
            // (right), which is what real PowerPoint renders. Only an explicit
            // val overrides this.
            var posRaw = legend.GetFirstChild<C.LegendPosition>()?.Val?.HasValue == true
                ? legend.GetFirstChild<C.LegendPosition>()!.Val!.InnerText : "r";
            node.Format["legend"] = posRaw switch
            {
                "b" => "bottom",
                "t" => "top",
                "l" => "left",
                "r" => "right",
                "tr" => "topRight",
                _ => posRaw
            };
            // Per-entry legend deletion (<c:legendEntry><c:idx/><c:delete val="1"/>).
            // The Setter writes these (legendEntry{N}.delete) but the Reader had
            // no read site, so the Get-driven dump silently reverted a hidden
            // legend entry on round-trip. Key uses 1-based ordinal = source idx+1,
            // matching the Setter's legendEntry{N} parse (N → idx N-1).
            foreach (var le in legend.Elements<C.LegendEntry>())
            {
                var leIdx = le.GetFirstChild<C.Index>()?.Val;
                var leDel = le.GetFirstChild<C.Delete>()?.Val;
                if (leDel?.HasValue == true && leDel.Value && leIdx?.HasValue == true)
                    node.Format[$"legendEntry{leIdx.Value + 1}.delete"] = "true";
            }
        }
        else
        {
            // Builder defaults to legend=bottom when prop absent; emit explicit
            // "none" so dump→replay round-trip preserves the no-legend state.
            node.Format["legend"] = "none";
        }

        // Chart-level dLbls lives as a direct child of the chart-group element
        // (c:barChart, c:lineChart, ...). Using Descendants pulled the first
        // series-level <c:dLbls> instead when it appeared earlier in XML order,
        // causing chart-level labelFont readback to mirror series 1's font.
        var labelGroup = plotArea.ChildElements
            .OfType<OpenXmlCompositeElement>()
            .Where(e => e is C.BarChart || e is C.LineChart || e is C.PieChart
                || e is C.AreaChart || e is C.Area3DChart || e is C.ScatterChart
                || e is C.DoughnutChart || e is C.Bar3DChart || e is C.Line3DChart
                || e is C.Pie3DChart || e is C.OfPieChart || e is C.BubbleChart
                || e is C.RadarChart || e is C.StockChart)
            .FirstOrDefault(g => g.GetFirstChild<C.DataLabels>() != null);
        var dataLabels = labelGroup?.GetFirstChild<C.DataLabels>();
        if (dataLabels != null)
        {
            var parts = new List<string>();
            if (dataLabels.GetFirstChild<C.ShowValue>()?.Val?.Value == true) parts.Add("value");
            if (dataLabels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true) parts.Add("category");
            if (dataLabels.GetFirstChild<C.ShowSeriesName>()?.Val?.Value == true) parts.Add("series");
            if (dataLabels.GetFirstChild<C.ShowPercent>()?.Val?.Value == true) parts.Add("percent");
            if (parts.Count > 0) node.Format["dataLabels"] = string.Join(",", parts);
            var dlPos = dataLabels.GetFirstChild<C.DataLabelPosition>()?.Val;
            if (dlPos?.HasValue == true)
            {
                // Return the schema-legal value verbatim (ctr, t, b, l, r,
                // outEnd, inEnd, inBase, bestFit). Stacked bar/column groupings
                // restrict dLblPos to {ctr, inBase, inEnd}; surfacing the raw
                // value lets callers verify exactly what was written and lines
                // up with our canonical-value rule (Get returns truth, Set
                // accepts friendly aliases). Friendly forms like "insideEnd"
                // remain accepted on the Set side via the alias map.
                //
                // ST_DLblPosPie restricts pie/pie3D to {bestFit, ctr, inEnd,
                // inBase}. A pie chart can still carry a stored outEnd/t/b/l/r
                // (Word silently treats it as bestFit), but emitting that value
                // would make a dump→batch replay reject the whole `add chart`
                // op and drop the chart. Suppress the invalid-for-pie value on
                // dump so the chart round-trips; the position is non-semantic
                // for pie anyway.
                var posText = dlPos.InnerText;
                var isPieGroup = labelGroup is C.PieChart or C.Pie3DChart;
                var pieValid = posText is "bestFit" or "ctr" or "inEnd" or "inBase";
                if (!isPieGroup || pieValid)
                    node.Format["labelPos"] = posText;
            }
        }

        // Chart style
        var style = chart.Parent?.GetFirstChild<C.Style>();
        if (style?.Val?.HasValue == true) node.Format["style"] = (int)style.Val.Value;

        // ManualLayout readback: plotArea, title, legend, trendlineLabel, displayUnitsLabel
        ReadManualLayout(plotArea, node, "plotArea");
        if (titleEl != null) ReadManualLayout(titleEl, node, "title");
        if (legend != null) ReadManualLayout(legend, node, "legend");
        var trendlineLbl = plotArea.Descendants<C.TrendlineLabel>().FirstOrDefault();
        if (trendlineLbl != null) ReadManualLayout(trendlineLbl, node, "trendlineLabel");
        var dispUnitsLbl = chart.Descendants<C.DisplayUnitsLabel>().FirstOrDefault();
        if (dispUnitsLbl != null) ReadManualLayout(dispUnitsLbl, node, "displayUnitsLabel");

        // Individual data label (dLbl) layout readback — first series
        var firstSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .FirstOrDefault(e => e.LocalName == "ser");
        var dLbls = firstSer?.GetFirstChild<C.DataLabels>();
        if (dLbls != null)
        {
            foreach (var dLbl in dLbls.Elements<C.DataLabel>())
            {
                var idx = dLbl.Index?.Val?.Value;
                if (idx == null) continue;
                var prefix = $"dataLabel{idx.Value + 1}";
                ReadManualLayout(dLbl, node, prefix);
                // Custom text
                var chartText = dLbl.GetFirstChild<C.ChartText>();
                var richText = chartText?.GetFirstChild<C.RichText>();
                var customText = richText?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
                if (customText != null) node.Format[$"{prefix}.text"] = customText;
                // Delete flag
                var delFlag = dLbl.GetFirstChild<C.Delete>()?.Val;
                if (delFlag?.HasValue == true && delFlag.Value) node.Format[$"{prefix}.delete"] = "true";
            }
        }

        // Plot area fill (plotArea uses C.ShapeProperties, not C.ChartShapeProperties)
        // R62 bt-2: pre-fix only emitted plotFill for <a:solidFill>; <a:gradFill>,
        // <a:noFill>, <a:pattFill>, <a:blipFill> silently dropped on dump→batch
        // replay. ReadFillSpec recognises all five; Setter side already routes
        // solid / gradient / "none" through BuildFillElement.
        var plotSpPr = plotArea.GetFirstChild<C.ShapeProperties>();
        var plotFillSpec = ReadFillSpec(plotSpPr);
        if (plotFillSpec != null) node.Format["plotFill"] = plotFillSpec;

        // Chart area fill (ChartSpace > spPr, NOT PlotArea)
        // Note: The SDK serializes ChartShapeProperties but deserializes it as C.ShapeProperties
        // after round-trip. Check both types, plus in-memory ChartShapeProperties.
        // R62 bt-2 (symmetric): pre-fix only recognised solidFill, so chartArea
        // gradient / no-fill round-tripped to nothing. ReadFillSpec covers all
        // five fill children.
        {
            OpenXmlCompositeElement? csSpPr = chart.Parent?.GetFirstChild<C.ShapeProperties>();
            if (csSpPr == null || ReadFillSpec(csSpPr) == null)
            {
                var csCSpPr = chart.Parent?.GetFirstChild<C.ChartShapeProperties>();
                if (csCSpPr != null) csSpPr = csCSpPr;
            }
            var chartFillSpec = ReadFillSpec(csSpPr);
            if (chartFillSpec != null) node.Format["chartFill"] = chartFillSpec;
        }

        // Gridlines: "true" for presence, detail in gridlineColor/gridlineWidth/gridlineDash
        var valAxisForGrid = plotArea.GetFirstChild<C.ValueAxis>();
        var majorGL = valAxisForGrid?.GetFirstChild<C.MajorGridlines>();
        if (majorGL != null)
        {
            node.Format["gridlines"] = "true";
            ReadGridlineDetail(majorGL, node, "gridline");
            // BUG-DUMP: the granular gridlineColor reads only the solidFill's
            // base scheme/rgb value and drops a:lumMod/a:lumOff (a tx1 gridline
            // tinted to 85% light gray rebuilt as solid black) plus the line's
            // cap/cmpd/algn/join. Capture the gridline <c:spPr> verbatim — same
            // approach as valAx.spPr — so the tint and line geometry round-trip.
            var majorGLSpPr = GetSpPrChildXml(majorGL);
            if (majorGLSpPr != null) node.InternalFormat["gridline.spPr"] = majorGLSpPr;
        }
        else if (valAxisForGrid != null)
        {
            node.Format["gridlines"] = "false";
        }
        var minorGL = valAxisForGrid?.GetFirstChild<C.MinorGridlines>();
        if (minorGL != null)
        {
            node.Format["minorGridlines"] = "true";
            ReadGridlineDetail(minorGL, node, "minorGridline");
            var minorGLSpPr = GetSpPrChildXml(minorGL);
            if (minorGLSpPr != null) node.InternalFormat["minorGridline.spPr"] = minorGLSpPr;
        }

        // GapWidth / Overlap from bar/column chart
        var barChart = plotArea.GetFirstChild<C.BarChart>();
        var gapWidthEl = barChart?.GetFirstChild<C.GapWidth>();
        if (gapWidthEl?.Val?.HasValue == true) node.Format["gapwidth"] = gapWidthEl.Val.Value.ToString();
        var overlapEl = barChart?.GetFirstChild<C.Overlap>();
        if (overlapEl?.Val?.HasValue == true) node.Format["overlap"] = overlapEl.Val.Value.ToString();
        // <c:serLines> on a stacked bar/column. Setter accepts
        // `serLines=true` via the "serlines"/"serieslines" case; without a
        // readback the source's element silently dropped on dump→replay.
        if (barChart?.GetFirstChild<C.SeriesLines>() != null)
            node.Format["serLines"] = "true";

        // gapDepth (3D bar/line/area chart depth spacing). Setter writes
        // <c:gapDepth> on the 3D chart group but the Reader had no read site,
        // so dump→replay dropped it. Only present on 3D charts, so Descendants
        // is unambiguous; emit only when the element exists (non-default).
        var gapDepthEl = plotArea.Descendants<C.GapDepth>().FirstOrDefault();
        if (gapDepthEl?.Val?.HasValue == true)
            node.Format["gapDepth"] = gapDepthEl.Val.Value.ToString();

        // Category-axis position (<c:catAx><c:axPos>). Setter's axisPosition
        // writes it on the category axis; the Reader never read it back, so a
        // non-default (top/left/right) tick position reverted to the bottom
        // default on dump→replay. axPos is a required element (always present),
        // so emit only when it differs from the category-axis default "b" to
        // avoid newly stamping axisPosition onto every existing chart's dump.
        var catAxisPos = plotArea.GetFirstChild<C.CategoryAxis>()
            ?.GetFirstChild<C.AxisPosition>()?.Val;
        if (catAxisPos?.HasValue == true && catAxisPos.InnerText != "b")
            node.Format["axisPosition"] = catAxisPos.InnerText switch
            {
                "t" => "top",
                "l" => "left",
                "r" => "right",
                _ => "bottom",
            };

        // CONSISTENCY(bar3d-shape): emit barShape so Set/Add shape=cone|cylinder|...
        // round-trips through Get. Lives on <c:bar3DChart><c:shape>.
        var bar3dForShape = plotArea.GetFirstChild<C.Bar3DChart>();
        var bar3dShape = bar3dForShape?.GetFirstChild<C.Shape>();
        if (bar3dShape?.Val?.HasValue == true)
        {
            var v = bar3dShape.Val.Value;
            string? barShapeStr = null;
            if (v == C.ShapeValues.Box) barShapeStr = "box";
            else if (v == C.ShapeValues.Cone) barShapeStr = "cone";
            else if (v == C.ShapeValues.ConeToMax) barShapeStr = "coneToMax";
            else if (v == C.ShapeValues.Cylinder) barShapeStr = "cylinder";
            else if (v == C.ShapeValues.Pyramid) barShapeStr = "pyramid";
            else if (v == C.ShapeValues.PyramidToMaximum) barShapeStr = "pyramidToMax";
            if (barShapeStr != null) node.Format["barShape"] = barShapeStr;
        }

        // Legend font (TextProperties on Legend element)
        if (legend != null)
        {
            var legendTp = legend.GetFirstChild<C.TextProperties>();
            if (legendTp != null)
            {
                var legendFontStr = ReadFontSpec(legendTp);
                if (legendFontStr != null) node.Format["legendFont"] = legendFontStr;
            }
        }

        // Axis font (TextProperties on value axis)
        var valAxisTp = valAxisForGrid?.GetFirstChild<C.TextProperties>();
        if (valAxisTp != null)
        {
            var axisFontStr = ReadFontSpec(valAxisTp);
            if (axisFontStr != null) node.Format["axisFont"] = axisFontStr;
        }

        // BUG-DUMP-R35-1: PER-AXIS text properties. The single `axisFont` key
        // above reads only the VALUE axis txPr; on rebuild that one font was
        // applied to BOTH axes, clobbering the category axis's distinct font
        // (e.g. catAx 10pt/666666 became valAx 9pt/999999). Capture each axis's
        // <c:txPr> verbatim (by local name, tolerant of the SDK post-reload
        // form) so per-axis fonts round-trip. The verbatim keys supersede the
        // lossy `axisFont` on replay (the emitter drops it). The category axis
        // may be a CategoryAxis or a DateAxis depending on catAxisType.
        var catAxisForTxPr = (OpenXmlElement?)plotArea.GetFirstChild<C.CategoryAxis>()
            ?? plotArea.GetFirstChild<C.DateAxis>();
        var catAxTxPrXml = GetTxPrChildXml(catAxisForTxPr);
        if (catAxTxPrXml != null) node.InternalFormat["catAx.txPr"] = catAxTxPrXml;
        var valAxTxPrXml = GetTxPrChildXml(valAxisForGrid);
        if (valAxTxPrXml != null) node.InternalFormat["valAx.txPr"] = valAxTxPrXml;

        // Secondary axis — emit the 1-based series indices bound to the
        // secondary axis so dump→replay round-trips. The Setter expects
        // "1,3" form (series indices); emitting bare "true" silently failed
        // parsing on replay because every comma-split token tried as int
        // produced [-1] then was filtered out.
        // R16-8: scatter and bubble charts inherently use two value axes
        // (X + Y), not a primary/secondary split. Reporting secondaryAxis for
        // them is a phantom readback that corrupts dump→replay. Skip them.
        var valAxes = plotArea.Elements<C.ValueAxis>().ToList();
        if (valAxes.Count > 1 && chartType is not ("scatter" or "bubble"))
        {
            // Map AxisId -> rank by document order; rank 0 = primary, 1 = secondary.
            var axisRank = new Dictionary<uint, int>();
            for (int ai = 0; ai < valAxes.Count; ai++)
            {
                var axId = SafeAxisIdVal(valAxes[ai].GetFirstChild<C.AxisId>());
                if (axId.HasValue) axisRank[axId.Value] = ai;
            }
            // Walk every series across every chart-type child of plotArea;
            // series indices are 1-based in document order matching how
            // ApplySecondaryAxis enumerates them.
            var secIdx = new List<int>();
            int seriesIdx = 0;
            foreach (var ct in plotArea.Elements<OpenXmlCompositeElement>())
            {
                foreach (var ser in ct.Elements<OpenXmlCompositeElement>()
                    .Where(e => e.LocalName == "ser"))
                {
                    seriesIdx++;
                    var seriesAxisIds = ser.Parent?.Elements<C.AxisId>().ToList()
                        ?? new List<C.AxisId>();
                    // A series's axis is determined by its parent chart-type
                    // element's c:axId children; primary vs secondary depends
                    // on which value-axis those IDs match.
                    var binds = seriesAxisIds
                        .Select(a => SafeAxisIdVal(a))
                        .Where(v => v.HasValue && axisRank.ContainsKey(v.Value))
                        .Select(v => axisRank[v!.Value]);
                    if (binds.Any(r => r >= 1)) secIdx.Add(seriesIdx);
                }
            }
            node.Format["secondaryAxis"] = secIdx.Count > 0
                ? string.Join(",", secIdx)
                : "true"; // Fallback only if we couldn't resolve any series.
        }

        // Axis label rotation (txPr/bodyPr/@rot in 60000ths of a degree)
        var catAxisForRot = (OpenXmlElement?)plotArea.GetFirstChild<C.CategoryAxis>()
            ?? plotArea.GetFirstChild<C.DateAxis>();
        var catAxisTxPr = catAxisForRot?.GetFirstChild<C.TextProperties>();
        var catAxisBodyPr = catAxisTxPr?.GetFirstChild<Drawing.BodyProperties>();
        if (catAxisBodyPr?.Rotation?.HasValue == true)
        {
            var deg = catAxisBodyPr.Rotation.Value / 60000.0;
            node.Format["xaxis.labelRotation"] = deg.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        }
        // fuzzer-3: scatter / bubble charts have no <c:catAx> — both axes are
        // <c:valAx>. Reading only the FIRST valAx as the y-axis silently drops
        // the x-axis labelRotation (and the symmetric case for the second
        // valAx as y on non-scatter charts already worked because it WAS the
        // first valAx). Disambiguate via the chart type: scatter/bubble place
        // the X axis FIRST among valAx siblings, all other chart types only
        // have one valAx and it's the Y axis.
        var valAxisList = plotArea.Elements<C.ValueAxis>().ToList();
        bool scatterLike = plotArea.GetFirstChild<C.ScatterChart>() != null
            || plotArea.GetFirstChild<C.BubbleChart>() != null;
        if (scatterLike && valAxisList.Count >= 1 && !node.Format.ContainsKey("xaxis.labelRotation"))
        {
            var xValAxBodyPr = valAxisList[0].GetFirstChild<C.TextProperties>()
                ?.GetFirstChild<Drawing.BodyProperties>();
            if (xValAxBodyPr?.Rotation?.HasValue == true)
            {
                var deg = xValAxBodyPr.Rotation.Value / 60000.0;
                node.Format["xaxis.labelRotation"] = deg.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
            }
        }
        var valAxisForY = scatterLike && valAxisList.Count >= 2
            ? valAxisList[1]
            : valAxisList.FirstOrDefault();
        var valAxisTxPrRot = valAxisForY?.GetFirstChild<C.TextProperties>();
        var valAxisBodyPr = valAxisTxPrRot?.GetFirstChild<Drawing.BodyProperties>();
        if (valAxisBodyPr?.Rotation?.HasValue == true)
        {
            var deg = valAxisBodyPr.Rotation.Value / 60000.0;
            node.Format["yaxis.labelRotation"] = deg.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        }

        // Axis titles. Capture the title paragraph's <a:pPr> verbatim alongside
        // the text: like the chart title, the axis-title builder (BuildChartTitle)
        // hard-codes the run to 14pt bold, so a source axis title sized only on
        // its defRPr (e.g. sz=800 b=0) rebuilt oversized and bold. The .pPr key
        // carries the source styling so the replay restores it.
        var valAxis = plotArea.GetFirstChild<C.ValueAxis>();
        var valAxisTitleEl = valAxis?.GetFirstChild<C.Title>();
        var valAxisTitle = valAxisTitleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (valAxisTitle != null)
        {
            node.Format["axisTitle"] = valAxisTitle;
            var vPPr = valAxisTitleEl!.Descendants<Drawing.ParagraphProperties>().FirstOrDefault();
            if (vPPr != null && AxisTitlePPrMeaningful(vPPr)) node.InternalFormat["axisTitle.pPr"] = vPPr.OuterXml;
        }

        var catAxis = plotArea.GetFirstChild<C.CategoryAxis>();
        var catAxisTitleEl = catAxis?.GetFirstChild<C.Title>();
        var catAxisTitle = catAxisTitleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (catAxisTitle != null)
        {
            node.Format["catTitle"] = catAxisTitle;
            var cPPr = catAxisTitleEl!.Descendants<Drawing.ParagraphProperties>().FirstOrDefault();
            if (cPPr != null && AxisTitlePPrMeaningful(cPPr)) node.InternalFormat["catTitle.pPr"] = cPPr.OuterXml;
        }

        // CONSISTENCY(cat-axis-type): emit the category-axis kind so Add/Set
        // can round-trip `catAxisType=date|category`. Default omitted when
        // a plain CategoryAxis is in use (matches the OOXML default).
        if (plotArea.GetFirstChild<C.DateAxis>() != null)
            node.Format["catAxisType"] = "date";

        // Axis scale
        var scaling = valAxis?.GetFirstChild<C.Scaling>();
        var minVal = scaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value;
        if (minVal != null) node.Format["axisMin"] = minVal;
        var maxVal = scaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value;
        if (maxVal != null) node.Format["axisMax"] = maxVal;

        var majorUnit = valAxis?.GetFirstChild<C.MajorUnit>()?.Val?.Value;
        if (majorUnit != null) node.Format["majorUnit"] = majorUnit;
        var minorUnit = valAxis?.GetFirstChild<C.MinorUnit>()?.Val?.Value;
        if (minorUnit != null) node.Format["minorUnit"] = minorUnit;

        var axisNumFmt = valAxis?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value;
        if (axisNumFmt != null && axisNumFmt != "General") node.Format["axisNumFmt"] = axisNumFmt;

        // Axis line styling
        var valAxisSpPr = valAxis?.GetFirstChild<C.ChartShapeProperties>();
        var valAxisOutline = valAxisSpPr?.GetFirstChild<Drawing.Outline>();
        if (valAxisOutline != null && valAxisOutline.GetFirstChild<Drawing.NoFill>() == null)
            ReadOutlineDetail(valAxisOutline, node, "valAxisLine");
        var catAxisSpPr = catAxis?.GetFirstChild<C.ChartShapeProperties>();
        var catAxisOutline = catAxisSpPr?.GetFirstChild<Drawing.Outline>();
        if (catAxisOutline != null && catAxisOutline.GetFirstChild<Drawing.NoFill>() == null)
            ReadOutlineDetail(catAxisOutline, node, "catAxisLine");

        // BUG-DUMP-R34-1: capture the value-axis line, category-axis line, and
        // plot-area border/fill as VERBATIM <c:spPr> OuterXml (same robust
        // approach as the R33 series/dPt/dLbls styling). The granular
        // valAxisLine/catAxisLine/plotArea.border readback above used the
        // strict typed C.ChartShapeProperties accessor, which the SDK
        // deserializes as the plain C.ShapeProperties form after a reload — so
        // the value-axis line and the plot-area border (a <a:ln> living
        // alongside the gradFill that `plotFill` already captures) were
        // silently dropped on dump→batch. GetSpPrChildXml reads the spPr child
        // by LOCAL NAME (works for both typed forms) and only emits when the
        // spPr carries meaningful styling (a:ln / fill / a:effectLst), so plain
        // charts stay clean. The verbatim key supersedes the lossy granular
        // keys on replay (the emitter drops them when the verbatim form is
        // present). These fragments reference theme colours (round-tripped via
        // the theme part) and carry NO external relationships.
        var valAxSpPrXml = GetSpPrChildXml(valAxis);
        if (valAxSpPrXml != null) node.InternalFormat["valAx.spPr"] = valAxSpPrXml;
        var catAxSpPrXml = GetSpPrChildXml(catAxis);
        if (catAxSpPrXml != null) node.InternalFormat["catAx.spPr"] = catAxSpPrXml;
        var plotAreaSpPrXml = GetSpPrChildXml(plotArea);
        if (plotAreaSpPrXml != null) node.InternalFormat["plotArea.spPr"] = plotAreaSpPrXml;

        // Axis visibility (c:delete)
        var valAxisDelete = valAxis?.GetFirstChild<C.Delete>();
        if (valAxisDelete?.Val?.HasValue == true && valAxisDelete.Val.Value)
            node.Format["valAxisVisible"] = "false";
        var catAxisDelete = catAxis?.GetFirstChild<C.Delete>();
        if (catAxisDelete?.Val?.HasValue == true && catAxisDelete.Val.Value)
            node.Format["catAxisVisible"] = "false";

        // Tick marks
        var valMajorTick = valAxis?.GetFirstChild<C.MajorTickMark>()?.Val;
        if (valMajorTick?.HasValue == true) node.Format["majorTickMark"] = valMajorTick.InnerText;
        var valMinorTick = valAxis?.GetFirstChild<C.MinorTickMark>()?.Val;
        if (valMinorTick?.HasValue == true) node.Format["minorTickMark"] = valMinorTick.InnerText;

        // Tick label position
        var valTickLblPos = valAxis?.GetFirstChild<C.TickLabelPosition>()?.Val;
        if (valTickLblPos?.HasValue == true) node.Format["tickLabelPos"] = valTickLblPos.InnerText;

        // Axis orientation
        var axisOrient = scaling?.GetFirstChild<C.Orientation>()?.Val;
        if (axisOrient?.HasValue == true && axisOrient.InnerText == "maxMin")
            node.Format["axisOrientation"] = "maxMin";

        // Log base
        var logBase = scaling?.GetFirstChild<C.LogBase>()?.Val?.Value;
        if (logBase != null) node.Format["logBase"] = logBase;

        // Display units
        var dispUnits = valAxis?.GetFirstChild<C.DisplayUnits>();
        var builtInUnit = dispUnits?.GetFirstChild<C.BuiltInUnit>()?.Val;
        if (builtInUnit?.HasValue == true) node.Format["dispUnits"] = builtInUnit.InnerText;

        // Crosses
        var crosses = valAxis?.GetFirstChild<C.Crosses>()?.Val;
        if (crosses?.HasValue == true) node.Format["crosses"] = crosses.InnerText;
        var crossesAt = valAxis?.GetFirstChild<C.CrossesAt>()?.Val?.Value;
        if (crossesAt != null) node.Format["crossesAt"] = crossesAt;
        var crossBetween = valAxis?.GetFirstChild<C.CrossBetween>()?.Val;
        if (crossBetween?.HasValue == true) node.Format["crossBetween"] = crossBetween.InnerText;

        // Category axis specifics
        var labelOffset = catAxis?.GetFirstChild<C.LabelOffset>()?.Val?.Value;
        if (labelOffset != null && labelOffset != 100) node.Format["labelOffset"] = labelOffset;
        var tickLblSkip = catAxis?.GetFirstChild<C.TickLabelSkip>()?.Val?.Value;
        if (tickLblSkip != null && tickLblSkip > 1) node.Format["tickLabelSkip"] = tickLblSkip;

        // Chart-level: smooth, showMarker, scatterStyle, varyColors, dispBlanksAs
        var lineChart = plotArea.GetFirstChild<C.LineChart>();
        var lineSmooth = lineChart?.GetFirstChild<C.Smooth>()?.Val;
        if (lineSmooth?.HasValue == true) node.Format["smooth"] = lineSmooth.Value ? "true" : "false";
        var showMarker = lineChart?.GetFirstChild<C.ShowMarker>()?.Val;
        if (showMarker?.HasValue == true) node.Format["showMarker"] = showMarker.Value ? "true" : "false";

        // Line-chart overlay elements: dropLines, hiLowLines, upDownBars.
        // Serialize back into Setter-spec form so dump→replay round-trips
        // visually (R31-F1: charts-line p7 lost the up/down bar overlay
        // because Reader was silent on these CT_LineChart children).
        if (lineChart != null)
        {
            var dropLinesEl = lineChart.GetFirstChild<C.DropLines>();
            if (dropLinesEl != null)
                node.Format["droplines"] = FormatLineOverlaySpec(dropLinesEl);

            var hiLowLinesEl = lineChart.GetFirstChild<C.HighLowLines>();
            if (hiLowLinesEl != null)
                node.Format["hilowlines"] = FormatLineOverlaySpec(hiLowLinesEl);

            var upDownBarsEl = lineChart.GetFirstChild<C.UpDownBars>();
            if (upDownBarsEl != null)
                node.Format["updownbars"] = FormatUpDownBarsSpec(upDownBarsEl);
        }

        var scatterChart = plotArea.GetFirstChild<C.ScatterChart>();
        var scatterStyle = scatterChart?.GetFirstChild<C.ScatterStyle>()?.Val;
        if (scatterStyle?.HasValue == true) node.Format["scatterStyle"] = scatterStyle.InnerText;

        var radarChart = plotArea.GetFirstChild<C.RadarChart>();
        var radarStyle = radarChart?.GetFirstChild<C.RadarStyle>()?.Val;
        if (radarStyle?.HasValue == true) node.Format["radarStyle"] = radarStyle.InnerText;

        var dispBlanksAs = chart.GetFirstChild<C.DisplayBlanksAs>()?.Val;
        if (dispBlanksAs?.HasValue == true) node.Format["dispBlanksAs"] = dispBlanksAs.InnerText;

        // varyColors: lives on the per-chart-type element (PieChart, BarChart, etc.).
        // Set writes the same value to every chart-type child of plotArea, so any
        // child carrying VaryColors faithfully represents the user-visible state.
        var varyColorsEl = plotArea.ChildElements
            .OfType<OpenXmlCompositeElement>()
            .Where(e => e.LocalName.Contains("Chart") || e.LocalName.Contains("chart"))
            .Select(ct => ct.GetFirstChild<C.VaryColors>())
            .FirstOrDefault(v => v?.Val?.HasValue == true);
        if (varyColorsEl?.Val?.HasValue == true)
            node.Format["varyColors"] = varyColorsEl.Val.Value ? "true" : "false";

        // roundedCorners
        var roundedCorners = chart.Parent?.GetFirstChild<C.RoundedCorners>()?.Val;
        if (roundedCorners?.HasValue == true) node.Format["roundedCorners"] = roundedCorners.Value ? "true" : "false";

        // View3D
        var view3d = chart.GetFirstChild<C.View3D>();
        if (view3d != null)
        {
            var rotX = view3d.GetFirstChild<C.RotateX>()?.Val?.Value;
            var rotY = view3d.GetFirstChild<C.RotateY>()?.Val?.Value;
            var persp = view3d.GetFirstChild<C.Perspective>()?.Val?.Value;
            var v3dParts = new List<string>();
            // Emit empty slot for missing child to preserve "not set" through
            // dump→replay. "0" placeholders caused Setter to write explicit
            // rotX/rotY/perspective=0 elements that PPT then renders as a flat
            // 3D camera (phantom rotation).
            v3dParts.Add(rotX != null ? rotX.Value.ToString() : "");
            v3dParts.Add(rotY != null ? rotY.Value.ToString() : "");
            v3dParts.Add(persp != null ? persp.Value.ToString() : "");
            // Suppress wholly-empty tuple (no children present at all).
            if (rotX != null || rotY != null || persp != null)
                node.Format["view3d"] = string.Join(",", v3dParts);
            if (rotX != null) node.Format["view3d.rotateX"] = (int)rotX.Value;
            if (rotY != null) node.Format["view3d.rotateY"] = (int)rotY.Value;
            if (persp != null) node.Format["view3d.perspective"] = (int)persp.Value;
        }

        // Data table
        var dataTable = plotArea.GetFirstChild<C.DataTable>();
        if (dataTable != null) node.Format["dataTable"] = "true";

        // Legend overlay
        var legendOverlay = legend?.GetFirstChild<C.Overlay>()?.Val;
        if (legendOverlay?.HasValue == true && legendOverlay.Value) node.Format["legend.overlay"] = "true";

        // Plot area border
        var plotOutline = plotSpPr?.GetFirstChild<Drawing.Outline>();
        if (plotOutline != null) ReadOutlineDetail(plotOutline, node, "plotArea.border");

        // Chart area border
        {
            var csSpPr = chart.Parent?.GetFirstChild<C.ShapeProperties>();
            var csOutline = csSpPr?.GetFirstChild<Drawing.Outline>();
            if (csOutline == null)
            {
                var csCSpPr = chart.Parent?.GetFirstChild<C.ChartShapeProperties>();
                csOutline = csCSpPr?.GetFirstChild<Drawing.Outline>();
            }
            if (csOutline != null) ReadOutlineDetail(csOutline, node, "chartArea.border");
            // Verbatim chartSpace <c:spPr> — like plotArea.spPr, the granular
            // chartArea.border.color reads only the base scheme value and drops
            // a:lumMod/a:lumOff (a tx1 frame tinted light gray rebuilt black).
            // Capture the frame fill + border verbatim so the tint round-trips.
            var chartAreaSpPr = chart.Parent != null ? GetSpPrChildXml(chart.Parent) : null;
            if (chartAreaSpPr != null) node.InternalFormat["chartArea.spPr"] = chartAreaSpPr;
        }

        // Chart-type-specific
        var pieChart = plotArea.GetFirstChild<C.PieChart>();
        var doughnutChart = plotArea.GetFirstChild<C.DoughnutChart>();
        // R13: firstSliceAngle lives on both pie and doughnut. Read from whichever
        // chart type is present so a doughnut's firstSliceAngle round-trips (the
        // Setter now writes it to the doughnut too).
        var firstSliceAngle = pieChart?.GetFirstChild<C.FirstSliceAngle>()?.Val?.Value
            ?? doughnutChart?.GetFirstChild<C.FirstSliceAngle>()?.Val?.Value;
        if (firstSliceAngle != null && firstSliceAngle != 0) node.Format["firstSliceAngle"] = firstSliceAngle;

        var holeSize = doughnutChart?.GetFirstChild<C.HoleSize>()?.Val?.Value;
        // CONSISTENCY(chart-format-type): emit as string to match sister
        // numeric chart props (gapwidth, overlap, explosion, style…).
        if (holeSize != null) node.Format["holeSize"] = ((int)holeSize).ToString();

        // Chart-level explosion (pie/doughnut): the Setter writes c:explosion
        // to every series uniformly. Surface as a single chart-level value
        // when all series agree; otherwise leave to per-series read-out.
        if (pieChart != null || doughnutChart != null)
        {
            var pieLikeSeries = plotArea.Descendants<OpenXmlCompositeElement>()
                .Where(e => e.LocalName == "ser" && (e.Parent is C.PieChart || e.Parent is C.DoughnutChart || e.Parent is C.Pie3DChart || e.Parent is C.OfPieChart))
                .ToList();
            if (pieLikeSeries.Count > 0)
            {
                uint? uniform = null;
                bool allSame = true;
                foreach (var ser in pieLikeSeries)
                {
                    var ex = ser.GetFirstChild<C.Explosion>()?.Val?.Value;
                    if (uniform == null) uniform = ex ?? 0;
                    else if ((ex ?? 0) != uniform) { allSame = false; break; }
                }
                if (allSame && uniform != null && uniform > 0)
                    node.Format["explosion"] = uniform.Value.ToString();
            }
        }

        var bubbleChart = plotArea.GetFirstChild<C.BubbleChart>();
        var bubbleScale = bubbleChart?.GetFirstChild<C.BubbleScale>()?.Val?.Value;
        if (bubbleScale != null && bubbleScale != 100) node.Format["bubbleScale"] = (int)bubbleScale;
        // fuzzer-1: <c:sizeRepresents val="width|area"> — controls how
        // bubble z-values map to bubble radius. PowerPoint defaults to
        // "area" when the element is absent. Emit only when explicitly
        // "width" so the default round-trip stays the empty element shape.
        var sizeReprVal = bubbleChart?.GetFirstChild<C.SizeRepresents>()?.Val?.Value;
        if (sizeReprVal != null && sizeReprVal == C.SizeRepresentsValues.Width)
            node.Format["sizeRepresents"] = "width";
        // fuzzer-2: <c:showNegBubbles val="0|1"> — controls visibility of
        // bubbles whose z-value is negative. Schema default is true; emit
        // only when explicitly false so untouched charts stay clean.
        var showNegVal = bubbleChart?.GetFirstChild<C.ShowNegativeBubbles>()?.Val?.Value;
        if (showNegVal == false) node.Format["showNegBubbles"] = "false";

        // DataLabels additional detail
        if (dataLabels != null)
        {
            var separator = dataLabels.GetFirstChild<C.Separator>()?.Text;
            if (separator != null) node.Format["dataLabels.separator"] = separator;
            var dlNumFmt = dataLabels.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value;
            if (dlNumFmt != null) node.Format["dataLabels.numFmt"] = dlNumFmt;

            // labelFont readback (R35-F3): BuildLabelTextProperties writes
            // <c:txPr><a:p><a:pPr><a:defRPr sz="..." b="..."><a:solidFill>
            // <a:srgbClr val="..."/></a:solidFill><a:latin typeface="..."/>
            // </a:defRPr></a:pPr></a:p></c:txPr>. Surface size / color / font
            // as dotted keys so dump→replay can rebuild the same spec via
            // labelFont=size:color:fontname.
            var dlDefRp = dataLabels.GetFirstChild<C.TextProperties>()
                ?.GetFirstChild<Drawing.Paragraph>()
                ?.GetFirstChild<Drawing.ParagraphProperties>()
                ?.GetFirstChild<Drawing.DefaultRunProperties>();
            if (dlDefRp != null)
            {
                if (dlDefRp.FontSize?.HasValue == true)
                    // CONSISTENCY(canonical-units / the project conventions): font
                    // sizes emit pt-qualified ("12pt"). Round-trip via labelFont.size
                    // accepts both "12" and "12pt" on input.
                    node.Format["labelFont.size"] = $"{dlDefRp.FontSize.Value / 100}pt";
                if (dlDefRp.Bold?.HasValue == true && dlDefRp.Bold.Value)
                    node.Format["labelFont.bold"] = "true";
                var dlLabelFill = dlDefRp.GetFirstChild<Drawing.SolidFill>();
                if (dlLabelFill != null)
                {
                    var dlLabelColor = ReadColorFromFill(dlLabelFill);
                    if (dlLabelColor != null)
                        node.Format["labelFont.color"] = dlLabelColor;
                }
                var dlLatin = dlDefRp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
                if (!string.IsNullOrEmpty(dlLatin))
                    node.Format["labelFont.name"] = dlLatin;
            }
        }

        var seriesCount = CountSeries(plotArea);
        node.Format["seriesCount"] = seriesCount;

        // R46 Major-3: emit per-series readback as `series{N}` = "Name:v1,v2,..."
        // so a `data=` Add round-trips (and pie-chart coalesce's data points are
        // visible). Mirrors the Setter input form (legacy series{N}=Name:1,2,3),
        // so dump→replay through the same key works for all single/multi series
        // charts. Reference-line overlay series are skipped (structural, not user).
        var seriesForReadback = ReadAllSeries(plotArea);
        var refMask = ReadReferenceLineMask(plotArea);
        int emittedIdx = 0;
        for (int si = 0; si < seriesForReadback.Count; si++)
        {
            if (si < refMask.Count && refMask[si]) continue;
            emittedIdx++;
            var (sName, sVals) = seriesForReadback[si];
            var vJoined = string.Join(",", sVals.Select(v =>
                v.ToString("G", System.Globalization.CultureInfo.InvariantCulture)));
            node.Format[$"series{emittedIdx}"] = string.IsNullOrEmpty(sName) || sName == "?"
                ? vJoined
                : $"{sName}:{vJoined}";
        }

        // Chart-level aggregate readback for series-level fan-out properties.
        // chart Set ('gradient' / 'marker') applies to every series — surface
        // the corresponding chart-level keys so a get-after-set round-trips
        // (schema declares gradient/marker get:true on chart-scope).
        var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();
        // R24 — emit the chart-level gradient as the same spec form the Setter
        // accepts ("colorA-colorB[:angle]") so dump→replay round-trips. Reading
        // the first series's GradientFill is sufficient because chart-scope
        // Set fans the same spec to every series (line 853 in Setter).
        var firstGradFill = allSer
            .Select(s => s.GetFirstChild<C.ChartShapeProperties>()?.GetFirstChild<Drawing.GradientFill>())
            .FirstOrDefault(g => g != null);
        if (firstGradFill != null)
        {
            var spec = ReadGradientSpec(firstGradFill);
            node.Format["gradient"] = spec ?? "true";
        }
        // Skip reference-line overlay series — their marker (val=none) is a
        // structural side-effect of AddReferenceLine, not a user-set marker.
        // Including them caused chart-level `marker=none` to be emitted on
        // any chart whose first real series had no explicit marker, then
        // dump→replay applied marker=none to series 1.
        var firstRealMarker = allSer
            .Where(s => !IsReferenceLineSeries(s))
            .Select(s => s.GetFirstChild<C.Marker>())
            .FirstOrDefault(m => m?.GetFirstChild<C.Symbol>()?.Val?.HasValue == true);
        var firstMarkerSym = firstRealMarker?.GetFirstChild<C.Symbol>()?.Val;
        if (firstMarkerSym != null) node.Format["marker"] = firstMarkerSym.InnerText;
        // markerColor fan-out: Setter accepts `marker=symbol:size:color`
        // but tests assert the bare-symbol readback, so emit color on a
        // separate key (mirrors markerSize). Reads the marker's spPr/solidFill.
        var firstMarkerFill = firstRealMarker?.GetFirstChild<C.ChartShapeProperties>()
            ?.GetFirstChild<Drawing.SolidFill>();
        if (firstMarkerFill != null)
        {
            var mColor = ReadColorFromFill(firstMarkerFill);
            if (mColor != null) node.Format["markerColor"] = mColor;
        }

        var cats = ReadCategories(plotArea);
        if (cats != null) node.Format["categories"] = string.Join(",", cats);

        var catsRef = ReadCategoriesRef(plotArea);
        if (catsRef != null) node.Format["categoriesRef"] = catsRef;

        // Trendline summary at chart level — scan first series with trendline
        var firstTrendlineSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser")
            .FirstOrDefault(s => s.GetFirstChild<C.Trendline>() != null);
        if (firstTrendlineSer != null)
        {
            var firstTl = firstTrendlineSer.GetFirstChild<C.Trendline>();
            var tlType = firstTl?.GetFirstChild<C.TrendlineType>()?.Val;
            if (tlType?.HasValue == true)
                node.Format["trendline"] = FormatTrendlineSpec(firstTl!, tlType.InnerText ?? "");
        }

        if (depth > 0)
        {
            var seriesList = ReadAllSeries(plotArea);
            for (int i = 0; i < seriesList.Count; i++)
            {
                var (sName, sValues) = seriesList[i];
                var seriesNode = new DocumentNode
                {
                    Path = $"{node.Path}/series[{i + 1}]",
                    Type = "series",
                    Text = sName
                };
                seriesNode.Format["name"] = sName;
                seriesNode.Format["values"] = string.Join(",", sValues.Select(v => v.ToString("G")));

                var serEl = plotArea.Descendants<OpenXmlCompositeElement>()
                    .Where(e => e.LocalName == "ser").ElementAtOrDefault(i);

                // Flag reference-line overlay series so the batch emitter
                // knows to omit them from `data=...` (the chart-level
                // `referenceLine=spec` rebuilds them via AddReferenceLine).
                if (serEl != null && IsReferenceLineSeries(serEl))
                    seriesNode.Format["refLine"] = "true";

                // Source c:idx / c:order — PowerPoint keys the theme accent
                // cycle (and stack order) off these, not off document
                // position. A combo dump reorders series by chart-group, so
                // rebuilding with positional idx recolors every series.
                if (serEl != null)
                {
                    var srcIdx = serEl.Elements<C.Index>().FirstOrDefault()?.Val?.Value;
                    if (srcIdx != null && srcIdx.Value != (uint)i)
                        seriesNode.Format["seriesIdx"] = srcIdx.Value.ToString();
                    var srcOrder = serEl.Elements<C.Order>().FirstOrDefault()?.Val?.Value;
                    if (srcOrder != null && srcOrder.Value != (uint)i)
                        seriesNode.Format["seriesOrder"] = srcOrder.Value.ToString();
                }

                // Source series with no explicit color (no <c:spPr> at all,
                // OR an spPr that only sets geometry like <a:ln w=…> without
                // any fill) inherit their color from the theme accent cycle.
                // Flag it so a replay suppresses the DefaultSeriesColors
                // injection (which would pin a modern Office palette over
                // the deck's own theme colors). Consumed by
                // SeriesWithNoCapturedFill via series{N}.inheritFill.
                if (serEl != null)
                {
                    var serSpPrEl = serEl.GetFirstChild<C.ChartShapeProperties>();
                    bool hasExplicitColor = serSpPrEl != null
                        && serSpPrEl.Descendants().Any(d => d.LocalName
                            is "solidFill" or "gradFill" or "pattFill" or "blipFill" or "noFill");
                    if (!hasExplicitColor)
                        seriesNode.Format["inheritFill"] = "true";
                }

                // Cell reference formulas (for series with NumberReference/StringReference)
                if (serEl != null)
                {
                    var valRef = ReadFormulaRef(serEl.GetFirstChild<C.Values>());
                    // Scatter/bubble series carry their numeric magnitude in
                    // <c:yVal>, not <c:val>. Without surfacing the yVal ref the
                    // Excel emitter's "all series have valuesRef" gate fails and
                    // the chart replays as disconnected literals (bubbles pile
                    // at x=0). The Builder maps info.ValuesRef → yVal on replay.
                    valRef ??= ReadFormulaRef(serEl.GetFirstChild<C.YValues>());
                    if (valRef != null) seriesNode.Format["valuesRef"] = valRef;

                    // Source numCache formatCode (e.g. #,##0). Data labels
                    // with numFmt sourceLinked=1 render THIS format — losing
                    // it drops thousands separators ("220,000" → "220000").
                    var valCacheFmt = serEl.GetFirstChild<C.Values>()
                        ?.GetFirstChild<C.NumberReference>()
                        ?.GetFirstChild<C.NumberingCache>()
                        ?.GetFirstChild<C.FormatCode>()?.Text;
                    if (!string.IsNullOrEmpty(valCacheFmt) && valCacheFmt != "General")
                        seriesNode.Format["valuesNumFmt"] = valCacheFmt;
                    var catRef = ReadFormulaRef(serEl.GetFirstChild<C.CategoryAxisData>());
                    // Scatter/bubble store their X range under <c:xVal>, not
                    // <c:cat>. The Builder rewrites xVal from info.CategoriesRef
                    // on replay, so surface the xVal ref as categoriesRef to keep
                    // the X axis live-linked to the source cells.
                    catRef ??= ReadFormulaRef(serEl.GetFirstChild<C.XValues>());
                    if (catRef != null) seriesNode.Format["categoriesRef"] = catRef;

                    // Sparse series: the source numCache/numLit may hold
                    // points at non-contiguous idx positions (blank cells in
                    // the source range — e.g. pts at idx 1,2 of ptCount 4).
                    // ReadAllSeries compacts them, so a replay would place
                    // every value at idx 0..n-1 and shift the points left.
                    // Surface the dense zero-padded value list plus the blank
                    // index list (0-based) so the batch emitter can rebuild
                    // the exact point placement via series{N}._blankIndexes.
                    var sparseValEl = (OpenXmlCompositeElement?)serEl.GetFirstChild<C.Values>()
                        ?? serEl.GetFirstChild<C.YValues>();
                    var sparse = ReadSparseNumericData(sparseValEl);
                    if (sparse != null)
                    {
                        seriesNode.Format["values"] = string.Join(",",
                            sparse.Value.padded.Select(v => v.ToString("G",
                                System.Globalization.CultureInfo.InvariantCulture)));
                        seriesNode.Format["blankIndexes"] = string.Join(",", sparse.Value.blanks);
                    }

                    // R44 major-2: scatter series carry X data under <c:xVal>
                    // (not <c:cat>). ReadAllSeries returns only Y values; surface
                    // X here so series.Format["x"] round-trips for scatter charts.
                    var xValEl = serEl.Elements<OpenXmlCompositeElement>()
                        .FirstOrDefault(e => e.LocalName == "xVal");
                    if (xValEl != null)
                    {
                        var xVals = ReadNumericData(xValEl);
                        if (xVals != null && xVals.Length > 0)
                            seriesNode.Format["x"] = string.Join(",", xVals.Select(v => v.ToString("G")));
                    }

                    // R52 bt-3: bubble series carry per-point sizes under
                    // <c:bubbleSize>; the data may live in <c:numLit> (literal)
                    // OR <c:numRef> (external cell range with <c:numCache>).
                    // The Builder always writes numLit on AddChart bubble, so
                    // a source-authored numRef was silently erased on
                    // dump→replay — the replayed bubble chart came back with
                    // BubbleSize equal to the YValues (BuildBubbleChart's
                    // default), losing both the cell ref and the cached
                    // pixel-sized geometry. Surface both ref + cache values
                    // (parallel to valuesRef + values for the y-axis) so the
                    // batch emitter can rebuild either form.
                    var bubbleSizeEl = serEl.GetFirstChild<C.BubbleSize>();
                    if (bubbleSizeEl != null)
                    {
                        var sizeVals = ReadNumericData(bubbleSizeEl);
                        if (sizeVals != null && sizeVals.Length > 0)
                            seriesNode.Format["bubbleSize"] = string.Join(",", sizeVals.Select(v => v.ToString("G")));
                        var sizeRef = ReadFormulaRef(bubbleSizeEl);
                        if (!string.IsNullOrEmpty(sizeRef))
                            seriesNode.Format["bubbleSizeRef"] = sizeRef;
                    }
                    var nameRefF = serEl.GetFirstChild<C.SeriesText>()
                        ?.GetFirstChild<C.StringReference>()
                        ?.GetFirstChild<C.Formula>()?.Text;
                    if (!string.IsNullOrEmpty(nameRefF)) seriesNode.Format["nameRef"] = nameRefF;
                }

                // BUG-DUMP-R33-1: capture per-series styling sub-elements as
                // VERBATIM OuterXml so dump→batch round-trips the full visual
                // fidelity (series <c:spPr> outline/shadow, every <c:dPt>
                // per-data-point override, and the rich <c:dLbls> num-format +
                // font). The granular attribute readback below (color, gradient,
                // outlineColor, shadow, point{N}.color, labelFont.*) only covers
                // a subset and silently flattens the rest. These three fragments
                // reference theme colours (round-tripped via the theme part) and
                // carry NO external relationships, so verbatim XML is safe. When
                // a verbatim key is present the emitter suppresses the matching
                // granular keys for that series so they don't double-apply.
                if (serEl != null)
                {
                    var rawSpPr = serEl.GetFirstChild<C.ChartShapeProperties>();
                    if (rawSpPr != null && HasMeaningfulStyling(rawSpPr))
                        seriesNode.InternalFormat["spPr"] = rawSpPr.OuterXml;

                    var rawDpts = serEl.Elements<C.DataPoint>()
                        .Where(dp => dp.GetFirstChild<C.ChartShapeProperties>() != null
                            || dp.GetFirstChild<C.Marker>() != null
                            || dp.GetFirstChild<C.Explosion>() != null)
                        .Select(dp => dp.OuterXml)
                        .ToList();
                    if (rawDpts.Count > 0)
                        // \x1e (record separator) joins the dPt fragments; it can
                        // never appear inside XML, so a literal split is safe.
                        // Stored in InternalFormat — verbatim OOXML is a dump→
                        // batch replay carrier, not user-facing Get output. The
                        // canonical per-point readback lives at point{N}.color/
                        // etc. below.
                        seriesNode.InternalFormat["dPt"] = string.Join("\x1e", rawDpts);

                    var rawDLbls = serEl.GetFirstChild<C.DataLabels>();
                    // Only round-trip dLbls verbatim when it carries rich styling
                    // (numFmt / spPr / txPr). A bare show-flag-only <c:dLbls> is
                    // already reconstructed by the existing dataLabels= readback,
                    // and replaying it verbatim would just duplicate that work.
                    // InternalFormat: same rationale as dPt above.
                    if (rawDLbls != null
                        && (rawDLbls.GetFirstChild<C.NumberingFormat>() != null
                            || rawDLbls.GetFirstChild<C.ChartShapeProperties>() != null
                            || rawDLbls.GetFirstChild<C.TextProperties>() != null))
                        seriesNode.InternalFormat["dLbls"] = rawDLbls.OuterXml;
                }

                var serSpPr = serEl?.GetFirstChild<C.ChartShapeProperties>();
                // NoFill round-trip: when ApplySeriesColor wrote <a:noFill/>
                // (color=none), Reader previously skipped emit and dump→replay
                // reverted to the default auto color. Surface "none" so the
                // setter side re-applies NoFill.
                if (serSpPr?.GetFirstChild<Drawing.NoFill>() != null)
                    seriesNode.Format["color"] = "none";
                // Line-based series (line/scatter/radar) carry their color on
                // the line stroke (<a:ln><a:solidFill>). Read it FIRST; fall
                // back to the bare <a:solidFill> for backward-compat with
                // files authored before the stroke-color fix.
                var serIsLineBased = serEl?.Parent?.LocalName
                    is "lineChart" or "scatterChart" or "radarChart";
                Drawing.SolidFill? serColor = null;
                if (serIsLineBased)
                    serColor = serSpPr?.GetFirstChild<Drawing.Outline>()?.GetFirstChild<Drawing.SolidFill>();
                serColor ??= serSpPr?.GetFirstChild<Drawing.SolidFill>();
                if (serColor != null)
                {
                    var colorVal = ReadColorFromFill(serColor);
                    if (colorVal != null) seriesNode.Format["color"] = colorVal;
                    // Alpha/transparency: schema declares both keys.
                    // - transparency is the percent-input mirror used on Add/Set
                    //   (100000 - alpha) / 1000 → 0..100 percent.
                    // - alpha is the raw OOXML units (0..100000 where 100000 =
                    //   opaque), schema-declared get:true and previously
                    //   not surfaced — meant Get readback hid the underlying
                    //   value when users set color with an alpha channel
                    //   (e.g. color=80FF0000).
                    var alphaEl = serColor.Descendants<Drawing.Alpha>().FirstOrDefault();
                    if (alphaEl?.Val?.HasValue == true)
                    {
                        var alphaUnits = (int)alphaEl.Val.Value;
                        seriesNode.Format["alpha"] = alphaUnits;
                        // transparency setter expects 0..100 percent — emit in
                        // the same unit so dump→batch round-trips cleanly.
                        // OOXML alpha is 0..100000 (100000 = fully opaque), so
                        // transparency% = (100000 - alpha) / 1000.
                        seriesNode.Format["transparency"] = Math.Round((100000 - alphaUnits) / 1000.0, 2);
                    }
                }
                // Gradient — emit the round-trippable spec form when possible.
                var gradFill = serSpPr?.GetFirstChild<Drawing.GradientFill>();
                if (gradFill != null)
                    seriesNode.Format["gradient"] = ReadGradientSpec(gradFill) ?? "true";
                // Line width
                var outline = serSpPr?.GetFirstChild<Drawing.Outline>();
                if (outline?.Width?.HasValue == true)
                    seriesNode.Format["lineWidth"] = Math.Round(outline.Width.Value / EmuConverter.EmuPerPointF, 2);
                // Line dash
                var prstDash = outline?.GetFirstChild<Drawing.PresetDash>();
                if (prstDash?.Val?.HasValue == true)
                    seriesNode.Format["lineDash"] = prstDash.Val.InnerText;
                // Outline color. For line-based series the stroke solidFill is
                // already surfaced as the series `color` above, so don't also
                // emit it as `outlineColor` (would double-encode the same value).
                var outlineFill = outline?.GetFirstChild<Drawing.SolidFill>();
                if (outlineFill != null && !serIsLineBased)
                {
                    var outColor = ReadColorFromFill(outlineFill);
                    if (outColor != null) seriesNode.Format["outlineColor"] = outColor;
                }
                // Shadow (from EffectList) — emit the full COLOR-BLUR-ANGLE-
                // DIST-OPACITY spec so dump→replay reconstructs the exact
                // <a:outerShdw> via the per-series shadow Setter (which
                // routes through DrawingEffectsHelper.BuildOuterShadow).
                // Prior emit of bare `shadow=true` silently dropped the
                // chart-series effect on round-trip (the per-series Setter
                // received "true" as the color spec and produced garbage).
                var effectList = serSpPr?.GetFirstChild<Drawing.EffectList>();
                var outerShadow = effectList?.GetFirstChild<Drawing.OuterShadow>();
                if (outerShadow != null)
                    seriesNode.Format["shadow"] = FormatOuterShadowSpec(outerShadow);
                // Marker
                var marker = serEl?.GetFirstChild<C.Marker>();
                var markerSymbol = marker?.GetFirstChild<C.Symbol>()?.Val;
                if (markerSymbol?.HasValue == true)
                    seriesNode.Format["marker"] = markerSymbol.InnerText;
                var markerSize = marker?.GetFirstChild<C.Size>()?.Val;
                if (markerSize?.HasValue == true)
                    seriesNode.Format["markerSize"] = (int)markerSize.Value;
                // markerColor: marker fill ships on its own key (see
                // chart-level fan-out above) so the bare `marker=symbol`
                // readback expected by tests is preserved.
                var markerFill = marker?.GetFirstChild<C.ChartShapeProperties>()
                    ?.GetFirstChild<Drawing.SolidFill>();
                if (markerFill != null)
                {
                    var mColor = ReadColorFromFill(markerFill);
                    if (mColor != null) seriesNode.Format["markerColor"] = mColor;
                }
                // Smooth
                var serSmooth = serEl?.GetFirstChild<C.Smooth>()?.Val;
                if (serSmooth?.HasValue == true) seriesNode.Format["smooth"] = serSmooth.Value ? "true" : "false";
                // Trendline(s): Excel allows multiple trendlines per series
                // (e.g. linear AND polynomial together). Emit all of them as
                // a semicolon-joined spec list so dump→replay re-applies each.
                // dispRSqr/dispEq mirror the FIRST trendline's display flags
                // (chart-level fan-out targets every trendline anyway).
                var trendlines = serEl?.Elements<C.Trendline>().ToList()
                    ?? new List<C.Trendline>();
                if (trendlines.Count > 0)
                {
                    var specs = new List<string>();
                    foreach (var tl in trendlines)
                    {
                        var tlType = tl.GetFirstChild<C.TrendlineType>()?.Val;
                        if (tlType?.HasValue == true)
                            specs.Add(FormatTrendlineSpec(tl, tlType.InnerText ?? ""));
                    }
                    if (specs.Count > 0)
                        seriesNode.Format["trendline"] = string.Join(";", specs);
                    var firstTl = trendlines[0];
                    var dispRSqr = firstTl.GetFirstChild<C.DisplayRSquaredValue>()?.Val;
                    if (dispRSqr?.HasValue == true && dispRSqr.Value) seriesNode.Format["trendline.dispRSqr"] = "true";
                    var dispEq = firstTl.GetFirstChild<C.DisplayEquation>()?.Val;
                    if (dispEq?.HasValue == true && dispEq.Value) seriesNode.Format["trendline.dispEq"] = "true";
                    // Forecast forward/backward + manual intercept. Previously
                    // unread, so dump→replay silently dropped them even though
                    // the Setter accepts trendline.forecastforward/backward/
                    // intercept. Emitted below via the chart-level fan-out keys.
                    var fwd = firstTl.GetFirstChild<C.Forward>()?.Val;
                    if (fwd?.HasValue == true) seriesNode.Format["trendline.forward"] = fwd.Value;
                    var bwd = firstTl.GetFirstChild<C.Backward>()?.Val;
                    if (bwd?.HasValue == true) seriesNode.Format["trendline.backward"] = bwd.Value;
                    var icpt = firstTl.GetFirstChild<C.Intercept>()?.Val;
                    if (icpt?.HasValue == true) seriesNode.Format["trendline.intercept"] = icpt.Value;
                    // CONSISTENCY(trendline-name-readback): the Setter writes
                    // a <c:trendlineLbl> with rich-text holding the user's
                    // name. Pull the text content back for Get parity.
                    var tlLblText = firstTl.GetFirstChild<C.TrendlineLabel>()
                        ?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
                    if (!string.IsNullOrEmpty(tlLblText))
                        seriesNode.Format["trendline.name"] = tlLblText;
                }
                // Error bars — emit as a "type:value" spec mirroring the
                // BuildErrorBars input form so dump→replay re-creates the
                // <c:errBars> element. Reading only the bare type lost the
                // magnitude (the <c:val>/<c:plus>/<c:minus> NumericLiteral),
                // and the per-series key `errBars` was also overshadowed by
                // chart-level errbars=... in batch emit.
                var errBars = serEl?.GetFirstChild<C.ErrorBars>();
                if (errBars != null)
                {
                    var errValType = errBars.GetFirstChild<C.ErrorBarValueType>()?.Val;
                    if (errValType?.HasValue == true)
                    {
                        var typeName = errValType.InnerText switch
                        {
                            "fixedVal" => "fixed",
                            "percentage" => "percent",
                            _ => errValType.InnerText  // OOXML camelCase: stdDev, stdErr, cust
                        };

                        // R55 bt-6: cust errBars carry per-direction value
                        // arrays under <c:plus>/<c:minus> via NumberReference
                        // (numCache) OR NumberLiteral (literal point list).
                        // The previous reader took only the first NumberLiteral
                        // point and collapsed multi-point cust down to a single
                        // magnitude, then the BuildErrorBars setter (with no
                        // "cust" arm) re-emitted it as fixedVal:0 — the entire
                        // cust spec was lost on dump-replay. Emit cust as
                        // "cust:<direction>:<plusCSV>:<minusCSV>" so Build can
                        // reconstruct both arrays.
                        if (typeName == "cust")
                        {
                            var direction = errBars.GetFirstChild<C.ErrorBarType>()?.Val?.InnerText
                                            ?? "both";
                            var plusCsv = ReadErrorBarSideCsv(errBars.GetFirstChild<C.Plus>());
                            var minusCsv = ReadErrorBarSideCsv(errBars.GetFirstChild<C.Minus>());
                            seriesNode.Format["errBars"] = $"cust:{direction}:{plusCsv}:{minusCsv}";
                        }
                        else
                        {
                            // Magnitude lives in either <c:val>, or shared
                            // <c:plus>/<c:minus> NumericLiteral first point.
                            string? mag = null;
                            var valEl = errBars.GetFirstChild<C.ErrorBarValue>()?.Val?.Value;
                            if (valEl.HasValue && valEl.Value != 0)
                                mag = valEl.Value.ToString("G",
                                    System.Globalization.CultureInfo.InvariantCulture);
                            else
                            {
                                var plusLit = errBars.GetFirstChild<C.Plus>()
                                    ?.GetFirstChild<C.NumberLiteral>();
                                var firstPt = plusLit?.Elements<C.NumericPoint>().FirstOrDefault();
                                var numStr = firstPt?.GetFirstChild<C.NumericValue>()?.Text;
                                if (!string.IsNullOrEmpty(numStr)) mag = numStr;
                            }
                            // Emit direction prefix when it's needed to keep
                            // round-trip lossless. plus/minus always carry the
                            // prefix (they're meaningful only with explicit
                            // direction). For ebDir=both, prefix when the type
                            // is the "bare direction" default (stdErr — the
                            // shape Setter produces from `errBars=both`), so
                            // Get round-trips "both:stdErr" back to the same
                            // input. Other types under direction=both stay
                            // implicit ("fixed:5") to preserve the prior
                            // R55-tested output for non-directional input.
                            // Always prefix the direction so Set/Get round-trip
                            // is lossless. R41's earlier guard `typeName ==
                            // "stdErr"` was too narrow: `both:fixed:5` and
                            // `both:percentage:10` carry ebDir=both with a
                            // non-stdErr type, and dropped the prefix on
                            // readback. Always emitting `both:` for explicit-
                            // direction Build paths keeps the form recoverable.
                            // (Sets without an explicit direction keyword
                            // still default to ebDir=both at Build time, so
                            // their readback also picks up the prefix — that's
                            // intentional now that R43 asserts the explicit
                            // round-trip form.)
                            var ebDir = errBars.GetFirstChild<C.ErrorBarType>()?.Val?.InnerText;
                            var dirPrefix = ebDir is "plus" or "minus" or "both"
                                ? ebDir + ":"
                                : "";
                            seriesNode.Format["errBars"] = mag != null
                                ? $"{dirPrefix}{typeName}:{mag}"
                                : $"{dirPrefix}{typeName}";
                        }
                    }
                }
                // InvertIfNegative
                var inv = serEl?.GetFirstChild<C.InvertIfNegative>()?.Val;
                if (inv?.HasValue == true && inv.Value) seriesNode.Format["invertIfNeg"] = "true";
                // Explosion (pie)
                var explosion = serEl?.GetFirstChild<C.Explosion>()?.Val?.Value;
                if (explosion != null && explosion > 0) seriesNode.Format["explosion"] = explosion;
                // Per-series labelFont readback. Mirrors the chart-level
                // labelFont readback above (line ~662) but scoped to this
                // series' own <c:dLbls> — Setter ApplySeriesLabelFont writes
                // here via series{N}.labelFont*=, and without per-series
                // readback dump→replay loses the spec.
                var serDLbls = serEl?.GetFirstChild<C.DataLabels>();
                // Per-series data-label SHOW flags (<c:showVal>, <c:showCatName>,
                // <c:showSerName>, <c:showPercent>). Without these the labels a
                // series displays (e.g. the value above each bar) were dropped on
                // dump→replay — only the labelFont styling below round-tripped.
                if (serDLbls != null)
                {
                    var dlFlags = new List<string>();
                    if (serDLbls.GetFirstChild<C.ShowValue>()?.Val?.Value == true) dlFlags.Add("value");
                    if (serDLbls.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true) dlFlags.Add("category");
                    if (serDLbls.GetFirstChild<C.ShowSeriesName>()?.Val?.Value == true) dlFlags.Add("series");
                    if (serDLbls.GetFirstChild<C.ShowPercent>()?.Val?.Value == true) dlFlags.Add("percent");
                    if (dlFlags.Count > 0)
                        seriesNode.Format["dataLabels"] = string.Join(",", dlFlags);

                    // Carry the whole <c:dLbls> verbatim whenever present —
                    // per-point dLbl, numFmt, separator, positions AND the
                    // plain show flags all ride it (the flag summary above is
                    // Get-friendly but was never replayed; a 3D bar's
                    // showVal=1 was silently dropped). The per-series
                    // `series{N}.dlbls` Set case re-inserts it in schema
                    // order.
                    seriesNode.Format["dlbls"] = serDLbls.OuterXml;
                }
                var serDlDefRp = serDLbls?.GetFirstChild<C.TextProperties>()
                    ?.GetFirstChild<Drawing.Paragraph>()
                    ?.GetFirstChild<Drawing.ParagraphProperties>()
                    ?.GetFirstChild<Drawing.DefaultRunProperties>();
                if (serDlDefRp != null)
                {
                    if (serDlDefRp.FontSize?.HasValue == true)
                        seriesNode.Format["labelFont.size"] = $"{serDlDefRp.FontSize.Value / 100}pt";
                    if (serDlDefRp.Bold?.HasValue == true && serDlDefRp.Bold.Value)
                        seriesNode.Format["labelFont.bold"] = "true";
                    var serDlLabelFill = serDlDefRp.GetFirstChild<Drawing.SolidFill>();
                    if (serDlLabelFill != null)
                    {
                        var serDlLabelColor = ReadColorFromFill(serDlLabelFill);
                        if (serDlLabelColor != null)
                            seriesNode.Format["labelFont.color"] = serDlLabelColor;
                    }
                    var serDlLatin = serDlDefRp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
                    if (!string.IsNullOrEmpty(serDlLatin))
                        seriesNode.Format["labelFont.name"] = serDlLatin;
                }
                // Data point colors + per-point marker / markerSize /
                // markerColor / explosion. Mirrors the series-level readback
                // above (markerSymbol/Size/spPr/SolidFill) so R38-B5 writer
                // output round-trips through dump→replay; without these the
                // dPt children are silently dropped on Get.
                if (serEl != null)
                {
                    foreach (var dPt in serEl.Elements<C.DataPoint>())
                    {
                        var ptIdx = dPt.Index?.Val?.Value;
                        if (ptIdx == null) continue;
                        var ptNum = ptIdx.Value + 1;
                        var ptSpPr = dPt.GetFirstChild<C.ChartShapeProperties>();
                        if (ptSpPr?.GetFirstChild<Drawing.NoFill>() != null)
                            seriesNode.Format[$"point{ptNum}.color"] = "none";
                        var ptFill = ptSpPr?.GetFirstChild<Drawing.SolidFill>();
                        if (ptFill != null)
                        {
                            var ptColor = ReadColorFromFill(ptFill);
                            if (ptColor != null) seriesNode.Format[$"point{ptNum}.color"] = ptColor;
                        }
                        // <c:marker> child of <c:dPt>
                        var ptMarker = dPt.GetFirstChild<C.Marker>();
                        if (ptMarker != null)
                        {
                            var ptMarkerSymbol = ptMarker.GetFirstChild<C.Symbol>()?.Val;
                            if (ptMarkerSymbol?.HasValue == true)
                                seriesNode.Format[$"point{ptNum}.marker"] = ptMarkerSymbol.InnerText;
                            var ptMarkerSize = ptMarker.GetFirstChild<C.Size>()?.Val;
                            if (ptMarkerSize?.HasValue == true)
                                seriesNode.Format[$"point{ptNum}.markerSize"] = (int)ptMarkerSize.Value;
                            var ptMarkerFill = ptMarker.GetFirstChild<C.ChartShapeProperties>()
                                ?.GetFirstChild<Drawing.SolidFill>();
                            if (ptMarkerFill != null)
                            {
                                var ptmColor = ReadColorFromFill(ptMarkerFill);
                                if (ptmColor != null) seriesNode.Format[$"point{ptNum}.markerColor"] = ptmColor;
                            }
                        }
                        // <c:explosion> child of <c:dPt> — pie/doughnut slice offset
                        var ptExplosion = dPt.GetFirstChild<C.Explosion>()?.Val;
                        if (ptExplosion?.HasValue == true)
                            seriesNode.Format[$"point{ptNum}.explosion"] = (int)ptExplosion.Value;
                    }
                }
                node.Children.Add(seriesNode);
            }
            node.ChildCount = seriesList.Count;
        }
        else
        {
            node.ChildCount = seriesCount;
        }
    }

    /// <summary>
    /// BUG-DUMP-R33-1: a series <c:spPr> is worth capturing verbatim only when
    /// it carries styling the granular <c:ser> readback can't fully reproduce —
    /// an outline (<a:ln>), an effect list (<a:outerShdw>/glow/…), or a gradient
    /// fill. A bare <c:spPr> holding just a <a:solidFill> (the common
    /// per-series colour) is already round-tripped by the `color` key, so
    /// emitting verbatim XML there would add noise and risk double-applying the
    /// fill. Keeping the verbatim path scoped to "rich" spPr means synthetic
    /// charts built by `add chart color=...` see no behavioural change.
    /// </summary>
    private static bool HasMeaningfulStyling(C.ChartShapeProperties spPr)
    {
        return spPr.GetFirstChild<Drawing.Outline>() != null
            || spPr.GetFirstChild<Drawing.EffectList>() != null
            || spPr.GetFirstChild<Drawing.GradientFill>() != null;
    }

    /// <summary>
    /// BUG-DUMP-R34-1: return the verbatim OuterXml of an element's <c:spPr>
    /// child when it carries meaningful styling — an outline (<a:ln>), any fill
    /// (<a:solidFill>/<a:gradFill>/<a:pattFill>/<a:blipFill>/<a:noFill>), or an
    /// effect list (<a:effectLst>). The spPr is located by LOCAL NAME so it
    /// works whether the SDK exposes it as the typed C.ChartShapeProperties
    /// (freshly built) or the plain C.ShapeProperties (after a part reload) —
    /// the strict typed accessors used elsewhere miss the reloaded form and
    /// silently drop the value-axis line / plot-area border. Returns null when
    /// there is no spPr or it holds nothing worth round-tripping verbatim, so
    /// plain charts emit no key.
    /// </summary>
    private static string? GetSpPrChildXml(OpenXmlElement? parent)
    {
        if (parent == null) return null;
        var spPr = parent.ChildElements
            .FirstOrDefault(e => e.LocalName == "spPr"
                && e.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/chart");
        if (spPr == null) return null;
        bool meaningful = spPr.ChildElements.Any(c =>
            c.LocalName is "ln" or "solidFill" or "gradFill" or "pattFill"
                or "blipFill" or "noFill" or "effectLst" or "effectDag");
        return meaningful ? spPr.OuterXml : null;
    }

    /// <summary>
    /// BUG-DUMP-R35-1: return the verbatim OuterXml of an element's <c:txPr>
    /// child when it carries meaningful text styling — a <a:defRPr> with a
    /// size / bold / fill / latin/ea/cs font, or a bodyPr rotation. The txPr is
    /// located by LOCAL NAME so it survives the SDK's post-reload form (same
    /// rationale as GetSpPrChildXml). Used to round-trip PER-AXIS fonts: the
    /// single legacy `axisFont` key reads only the value axis, so the category
    /// axis's distinct font was clobbered by the value axis's on rebuild.
    /// Returns null for an empty/styling-free txPr so plain charts emit nothing.
    /// </summary>
    private static string? GetTxPrChildXml(OpenXmlElement? parent)
    {
        if (parent == null) return null;
        var txPr = parent.ChildElements
            .FirstOrDefault(e => e.LocalName == "txPr"
                && e.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/chart");
        if (txPr == null) return null;
        // Meaningful when any defRPr carries explicit styling, or the bodyPr
        // sets a rotation/vert (rotation already round-trips via
        // xaxis/yaxis.labelRotation, but capturing the whole txPr verbatim is
        // harmless and keeps the font + rotation as one fragment).
        var defRPr = txPr.Descendants()
            .FirstOrDefault(e => e.LocalName == "defRPr");
        bool meaningfulFont = defRPr != null && (
            defRPr.GetAttributes().Any(a => a.LocalName is "sz" or "b" or "i" or "u")
            || defRPr.ChildElements.Any(c => c.LocalName is "solidFill" or "latin"
                or "ea" or "cs" or "highlight"));
        if (meaningfulFont) return txPr.OuterXml;
        return null;
    }

    /// <summary>
    /// R55 bt-6: read one side of a cust errBars (<c:plus> or <c:minus>) as a
    /// comma-separated value list. Prefers the numCache inside <c:numRef>
    /// (PowerPoint always writes the cached values alongside the formula);
    /// falls back to <c:numLit> (literal point list) when the side is bound
    /// to inline values rather than a sheet reference. Empty when no side
    /// element was supplied.
    /// </summary>
    private static string ReadErrorBarSideCsv(OpenXmlCompositeElement? side)
    {
        if (side == null) return "";
        // numRef.numCache first — numRef is the common cust authoring form.
        var numRef = side.GetFirstChild<C.NumberReference>();
        var cache = numRef?.NumberingCache;
        if (cache != null)
        {
            var ptCount = cache.Elements<C.NumericPoint>().Count();
            if (ptCount > 0)
            {
                return string.Join(",",
                    cache.Elements<C.NumericPoint>()
                        .OrderBy(p => p.Index?.Value ?? 0u)
                        .Select(p => p.GetFirstChild<C.NumericValue>()?.Text ?? ""));
            }
        }
        // Fallback to inline numLit.
        var numLit = side.GetFirstChild<C.NumberLiteral>();
        if (numLit != null)
        {
            return string.Join(",",
                numLit.Elements<C.NumericPoint>()
                    .OrderBy(p => p.Index?.Value ?? 0u)
                    .Select(p => p.GetFirstChild<C.NumericValue>()?.Text ?? ""));
        }
        return "";
    }

    internal static string? DetectChartType(C.PlotArea plotArea)
    {
        // Count real chart-type elements. A LineChart containing only reference-line-shaped
        // series (flat values, no marker, dashed outline) is a ref-line overlay added by
        // AddReferenceLine — it must not promote the underlying chart to a "combo".
        var chartTypeCount = plotArea.ChildElements
            .Count(e => (e is C.BarChart or C.LineChart or C.PieChart or C.AreaChart or C.Area3DChart
                or C.ScatterChart or C.DoughnutChart or C.Bar3DChart or C.Line3DChart or C.Pie3DChart
                or C.OfPieChart
                or C.BubbleChart or C.RadarChart or C.StockChart)
                && !(e is C.LineChart lc && IsReferenceLineOnlyChart(lc)));
        if (chartTypeCount > 1) return "combo";

        // The dispatch below picks the first real chart-type child. A
        // reference-line-only LineChart sibling (added by AddReferenceLine on
        // an area/bar/column chart) must not steal the dispatch -- otherwise
        // a chart authored as `type=area` + `referenceLine=60` reports
        // chartType=line on Get, and dump→replay rebuilds it as a single
        // lineChart with no areaChart in plotArea.
        bool IsRefOnly(OpenXmlElement el) => el is C.LineChart lc2 && IsReferenceLineOnlyChart(lc2);

        if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart bar)
        {
            var dir = bar.GetFirstChild<C.BarDirection>()?.Val?.Value;
            var grp = bar.GetFirstChild<C.BarGrouping>()?.Val?.InnerText;
            var prefix = dir == C.BarDirectionValues.Bar ? "bar" : "column";
            if (grp == "stacked")
            {
                // Detect waterfall chart: stacked bar with 3 series where first is "Base" with NoFill
                if (IsWaterfallPattern(bar))
                    return "waterfall";
                return $"{prefix}_stacked";
            }
            if (grp == "percentStacked") return $"{prefix}_percentStacked";
            return prefix;
        }
        if (plotArea.Elements<C.LineChart>().FirstOrDefault(lc => !IsRefOnly(lc)) is C.LineChart lineCh)
        {
            // Mirror bar/area: encode stacked / percentStacked into the
            // chartType token so dump→replay rebuilds the right grouping
            // (Builder reads chartType only — no separate `grouping` Set key).
            var lineGrp = lineCh.GetFirstChild<C.Grouping>()?.Val?.InnerText;
            if (lineGrp == "stacked") return "line_stacked";
            if (lineGrp == "percentStacked") return "line_percentStacked";
            return "line";
        }
        if (plotArea.GetFirstChild<C.PieChart>() != null) return "pie";
        if (plotArea.GetFirstChild<C.OfPieChart>() is C.OfPieChart ofPie)
        {
            // CT_OfPieChart distinguishes via c:ofPieType (pie | bar).
            var ofPieType = ofPie.GetFirstChild<C.OfPieType>()?.Val?.Value;
            return ofPieType == C.OfPieValues.Bar ? "barOfPie" : "pieOfPie";
        }
        if (plotArea.GetFirstChild<C.DoughnutChart>() != null) return "doughnut";
        if (plotArea.GetFirstChild<C.AreaChart>() is C.AreaChart area)
        {
            var areaGrp = area.GetFirstChild<C.Grouping>()?.Val?.InnerText;
            if (areaGrp == "stacked") return "area_stacked";
            if (areaGrp == "percentStacked") return "area_percentStacked";
            return "area";
        }
        if (plotArea.GetFirstChild<C.Area3DChart>() is C.Area3DChart area3d)
        {
            var area3dGrp = area3d.GetFirstChild<C.Grouping>()?.Val?.InnerText;
            if (area3dGrp == "stacked") return "area3d_stacked";
            if (area3dGrp == "percentStacked") return "area3d_percentStacked";
            return "area3d";
        }
        if (plotArea.GetFirstChild<C.ScatterChart>() != null) return "scatter";
        if (plotArea.GetFirstChild<C.BubbleChart>() != null) return "bubble";
        if (plotArea.GetFirstChild<C.RadarChart>() != null) return "radar";
        if (plotArea.GetFirstChild<C.StockChart>() != null) return "stock";
        if (plotArea.GetFirstChild<C.Bar3DChart>() is C.Bar3DChart bar3d)
        {
            var dir3d = bar3d.GetFirstChild<C.BarDirection>()?.Val?.Value;
            var grp3d = bar3d.GetFirstChild<C.BarGrouping>()?.Val?.InnerText;
            var prefix3d = dir3d == C.BarDirectionValues.Bar ? "bar" : "column";
            var suffix3d = grp3d == "stacked" ? "_stacked"
                : grp3d == "percentStacked" ? "_percentStacked"
                : "";
            return $"{prefix3d}3d{suffix3d}";
        }
        if (plotArea.GetFirstChild<C.Line3DChart>() != null) return "line3d";
        if (plotArea.GetFirstChild<C.Pie3DChart>() != null) return "pie3d";
        return null;
    }

    /// <summary>
    /// A reference-line series has (a) all values equal (flat horizontal line in OOXML terms),
    /// (b) marker set to None, and (c) outline with a preset dash style. This matches the
    /// shape that AddReferenceLine emits and is used to detect/remove overlays.
    /// </summary>
    internal static bool IsReferenceLineSeries(OpenXmlCompositeElement ser)
    {
        if (ser.LocalName != "ser") return false;

        var marker = ser.GetFirstChild<C.Marker>();
        if (marker?.GetFirstChild<C.Symbol>()?.Val?.Value != C.MarkerStyleValues.None) return false;

        var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
        var outline = spPr?.GetFirstChild<Drawing.Outline>();
        if (outline?.GetFirstChild<Drawing.PresetDash>() == null) return false;

        // Flat values — every NumericPoint has the same text. Must have at least 1 literal point.
        var numLit = ser.GetFirstChild<C.Values>()?.GetFirstChild<C.NumberLiteral>();
        if (numLit == null) return false;
        var distinct = numLit.Elements<C.NumericPoint>()
            .Select(p => p.InnerText)
            .Distinct()
            .Take(2)
            .Count();
        return distinct == 1;
    }

    /// <summary>
    /// True if a LineChart is made up entirely of reference-line series (i.e. it is a
    /// ref-line overlay, not a real line chart). Empty LineCharts do not count.
    /// </summary>
    internal static bool IsReferenceLineOnlyChart(C.LineChart lineChart)
    {
        var sers = lineChart.Elements<C.LineChartSeries>().ToList();
        if (sers.Count == 0) return false;
        return sers.All(IsReferenceLineSeries);
    }

    /// <summary>
    /// Read all reference-line overlays from a plot area. Returns value, label, color,
    /// line width in points, and dash style name. Colors come back as 6-digit hex without
    /// the '#' prefix; dash name is the OOXML PresetLineDashValues InnerText (e.g. "sysDash").
    /// </summary>
    internal static List<(string Name, double Value, string Color, double WidthPt, string Dash)> ReadReferenceLines(C.PlotArea plotArea, Dictionary<string, string>? themeColors = null)
    {
        var result = new List<(string, double, string, double, string)>();
        foreach (var lineChart in plotArea.Elements<C.LineChart>())
        {
            foreach (var ser in lineChart.Elements<C.LineChartSeries>())
            {
                if (!IsReferenceLineSeries(ser)) continue;

                // Value: any NumericPoint (all equal by definition of ref-line series)
                var numLit = ser.GetFirstChild<C.Values>()?.GetFirstChild<C.NumberLiteral>();
                var pt = numLit?.Elements<C.NumericPoint>().FirstOrDefault();
                if (pt == null) continue;
                if (!double.TryParse(pt.InnerText,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out var val))
                    continue;

                var name = ser.GetFirstChild<C.SeriesText>()
                    ?.Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? "";

                var outline = ser.GetFirstChild<C.ChartShapeProperties>()?.GetFirstChild<Drawing.Outline>();
                var widthEmu = outline?.Width?.Value ?? 19050;
                var widthPt = widthEmu / EmuConverter.EmuPerPointF;

                // Color: solidFill srgbClr hex, or schemeClr resolved through the theme map
                // (this secondary reader previously handled only srgbClr, so a themed reference
                // line fell back to the red default instead of its accent color).
                var color = "FF0000";
                var refFill = outline?.GetFirstChild<Drawing.SolidFill>();
                var srgb = refFill?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                if (!string.IsNullOrEmpty(srgb))
                    color = srgb;
                else if (refFill?.GetFirstChild<Drawing.SchemeColor>()?.Val?.InnerText is string schemeName
                         && themeColors != null)
                {
                    var canonical = ParseHelpers.NormalizeSchemeColorName(schemeName) ?? schemeName;
                    if (themeColors.TryGetValue(canonical, out var hex) || themeColors.TryGetValue(schemeName, out hex))
                        color = hex;
                }

                var dashVal = outline?.GetFirstChild<Drawing.PresetDash>()?.Val;
                var dash = dashVal?.InnerText ?? "dash";

                result.Add((name, val, color, widthPt, dash));
            }
        }
        return result;
    }

    /// <summary>
    /// Detect waterfall chart pattern: a stacked bar chart with exactly 3 series
    /// where the first series is named "Base" and has NoFill (invisible base).
    /// </summary>
    private static bool IsWaterfallPattern(C.BarChart bar)
    {
        var series = bar.Elements<C.BarChartSeries>().ToList();
        if (series.Count != 3) return false;

        // First series should be "Base" with NoFill
        var firstSerName = series[0].GetFirstChild<C.SeriesText>()
            ?.GetFirstChild<C.StringReference>()?.GetFirstChild<C.StringCache>()
            ?.GetFirstChild<C.StringPoint>()?.GetFirstChild<C.NumericValue>()?.Text
            ?? series[0].GetFirstChild<C.SeriesText>()
            ?.GetFirstChild<C.NumericValue>()?.Text;

        if (!string.Equals(firstSerName, "Base", StringComparison.OrdinalIgnoreCase))
            return false;

        // First series should have NoFill in its shape properties
        var baseSpPr = series[0].GetFirstChild<C.ChartShapeProperties>();
        if (baseSpPr?.GetFirstChild<Drawing.NoFill>() == null)
            return false;

        return true;
    }

    internal static int CountSeries(C.PlotArea plotArea)
    {
        return plotArea.Descendants<C.Index>()
            .Count(idx => idx.Parent?.LocalName == "ser");
    }

    internal static string[]? ReadCategories(C.PlotArea plotArea)
    {
        var catData = plotArea.Descendants<C.CategoryAxisData>().FirstOrDefault();
        if (catData == null)
        {
            // R44 major-2: scatter charts have no <c:cat>; their X-axis data
            // lives under each series' <c:xVal>. Fall back to the first ser's
            // xVal so chart-level Format["categories"] still surfaces the
            // X-axis labels for scatter dump→replay round-trip.
            var firstSer = plotArea.Descendants<OpenXmlCompositeElement>()
                .FirstOrDefault(e => e.LocalName == "ser" && e.Parent != null &&
                    (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart")));
            var xValEl = firstSer?.Elements<OpenXmlCompositeElement>()
                .FirstOrDefault(e => e.LocalName == "xVal");
            if (xValEl != null)
            {
                var xVals = ReadNumericData(xValEl);
                if (xVals != null && xVals.Length > 0)
                    return xVals.Select(v => v.ToString("G")).ToArray();
            }
            return null;
        }

        var strLit = catData.GetFirstChild<C.StringLiteral>();
        if (strLit != null)
        {
            return strLit.Elements<C.StringPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => p.GetFirstChild<C.NumericValue>()?.Text ?? "")
                .ToArray();
        }

        var strRef = catData.GetFirstChild<C.StringReference>();
        var strCache = strRef?.GetFirstChild<C.StringCache>();
        if (strCache != null)
        {
            return strCache.Elements<C.StringPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => p.GetFirstChild<C.NumericValue>()?.Text ?? "")
                .ToArray();
        }

        // Numeric category axes: Excel/PowerPoint emit <c:numRef>/<c:numLit> when
        // the category labels are numbers (years, quarters, etc). Without these
        // branches the axis labels render blank. Mirror ReadNumericData's chain.
        var numRef = catData.GetFirstChild<C.NumberReference>();
        var numCache = numRef?.GetFirstChild<C.NumberingCache>();
        if (numCache != null)
        {
            return numCache.Elements<C.NumericPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => p.GetFirstChild<C.NumericValue>()?.Text ?? "")
                .ToArray();
        }

        var numLit = catData.GetFirstChild<C.NumberLiteral>();
        if (numLit != null)
        {
            return numLit.Elements<C.NumericPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => p.GetFirstChild<C.NumericValue>()?.Text ?? "")
                .ToArray();
        }

        // Multi-level (grouped) category axis (<c:multiLvlStrRef>, emitted when Excel
        // grouped row/column headers feed the category axis). PowerPoint draws a
        // hierarchical axis; the FIRST <c:lvl> holds the innermost, per-point labels
        // nearest the axis (later lvls are coarser groupings). The single-level renderer
        // surfaces that innermost level (the grouping rows are a known limitation).
        // Without this branch all category labels rendered blank.
        var multiLvlRef = catData.Elements().FirstOrDefault(e => e.LocalName == "multiLvlStrRef");
        var multiCache = multiLvlRef?.Elements().FirstOrDefault(e => e.LocalName == "multiLvlStrCache");
        var firstLvl = multiCache?.Elements().FirstOrDefault(e => e.LocalName == "lvl");
        if (firstLvl != null)
        {
            return firstLvl.Elements().Where(e => e.LocalName == "pt")
                .OrderBy(p => uint.TryParse(
                    p.GetAttributes().FirstOrDefault(a => a.LocalName == "idx").Value, out var n) ? n : 0u)
                .Select(p => p.Elements().FirstOrDefault(e => e.LocalName == "v")?.InnerText ?? "")
                .ToArray();
        }

        // StringReference without cache — return null (data lives in cells)
        // The formula is read separately via ReadFormulaRef
        return null;
    }

    /// <summary>
    /// Read the categories formula reference from the first CategoryAxisData element.
    /// Returns null if no reference found (literal categories).
    /// </summary>
    internal static string? ReadCategoriesRef(C.PlotArea plotArea)
    {
        var catData = plotArea.Descendants<C.CategoryAxisData>().FirstOrDefault();
        return ReadFormulaRef(catData);
    }

    /// <summary>
    /// Read the first cached string value from a <c:strLit> child of <c:tx>.
    /// <c:strLit> is not a schema-valid child of <c:tx> (CT_SerTx allows only
    /// <c:strRef> or <c:v>), but PowerPoint accepts the form and authoring
    /// tools occasionally emit it. The SDK can't type those elements, so we
    /// walk by LocalName and pull the first <c:pt>/<c:v>.
    /// </summary>
    private static string? ReadStrLitFirstValue(OpenXmlElement? serText)
    {
        if (serText == null) return null;
        var strLit = serText.Elements().FirstOrDefault(e => e.LocalName == "strLit");
        if (strLit == null) return null;
        var pt = strLit.Elements().FirstOrDefault(e => e.LocalName == "pt");
        if (pt == null) return null;
        var v = pt.Elements().FirstOrDefault(e => e.LocalName == "v");
        return v?.InnerText;
    }

    internal static List<(string name, double[] values)> ReadAllSeries(C.PlotArea plotArea)
    {
        var result = new List<(string name, double[] values)>();

        foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent != null &&
                (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart"))))
        {
            var serText = ser.GetFirstChild<C.SeriesText>();
            // c:tx may carry <c:strRef> (cached cell value), bare <c:v> (literal),
            // or — non-schema but emitted by some authoring tools — <c:strLit>
            // (literal cached series-name table mirroring the category form).
            // PowerPoint renders all three; SDK only types strRef and bare <c:v>,
            // so the strLit branch is read via descendant element-name probing.
            // Prefer the cached value from strRef, fall back to the formula, then
            // literal <c:v>, so users who set series{N}.name=Sheet1!A1 still get
            // a meaningful name back from Get.
            string name = "?";
            var strRef = serText?.GetFirstChild<C.StringReference>();
            if (strRef != null)
            {
                var cached = strRef.GetFirstChild<C.StringCache>()
                    ?.GetFirstChild<C.StringPoint>()
                    ?.GetFirstChild<C.NumericValue>()?.Text;
                name = !string.IsNullOrEmpty(cached)
                    ? cached
                    : (strRef.GetFirstChild<C.Formula>()?.Text ?? "?");
            }
            else
            {
                name = serText?.Descendants<C.NumericValue>().FirstOrDefault()?.Text
                    ?? ReadStrLitFirstValue(serText)
                    ?? "?";
            }

            var values = ReadNumericData(ser.GetFirstChild<C.Values>())
                ?? ReadNumericData(ser.Elements<OpenXmlCompositeElement>()
                    .FirstOrDefault(e => e.LocalName == "yVal"))
                ?? Array.Empty<double>();

            result.Add((name, values));
        }

        return result;
    }

    /// <summary>
    /// Enumerate ser elements in the same order ReadAllSeries visits them, returning
    /// `true` for each series that is a reference-line overlay. The caller can zip
    /// this with the ReadAllSeries output to filter out ref-line entries without
    /// re-walking the OOXML tree.
    /// </summary>
    internal static List<bool> ReadReferenceLineMask(C.PlotArea plotArea)
    {
        var result = new List<bool>();
        foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent != null &&
                (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart"))))
        {
            result.Add(IsReferenceLineSeries(ser));
        }
        return result;
    }

    /// <summary>
    /// Detect a sparse numeric point list (fewer points than ptCount, i.e.
    /// blank source cells) inside a val/yVal's numCache or numLit. Returns
    /// the dense zero-padded value array plus the 0-based blank indexes, or
    /// null when the data is dense (or absent / malformed).
    /// </summary>
    internal static (double[] padded, int[] blanks)? ReadSparseNumericData(OpenXmlCompositeElement? valElement)
    {
        if (valElement == null) return null;
        var container = (OpenXmlCompositeElement?)valElement
                .GetFirstChild<C.NumberReference>()?.GetFirstChild<C.NumberingCache>()
            ?? valElement.GetFirstChild<C.NumberLiteral>();
        if (container == null) return null;
        var ptCount = (int?)container.GetFirstChild<C.PointCount>()?.Val?.Value ?? -1;
        var pts = container.Elements<C.NumericPoint>().ToList();
        if (ptCount <= 0 || pts.Count == 0 || pts.Count >= ptCount) return null;
        var byIdx = new Dictionary<int, double>();
        foreach (var pt in pts)
        {
            if (pt.Index?.Value is not uint uidx || uidx >= (uint)ptCount) return null;
            double.TryParse(pt.GetFirstChild<C.NumericValue>()?.Text,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v);
            byIdx[(int)uidx] = v;
        }
        var padded = new double[ptCount];
        var blanks = new List<int>();
        for (int i = 0; i < ptCount; i++)
        {
            if (byIdx.TryGetValue(i, out var v)) padded[i] = v;
            else blanks.Add(i);
        }
        return (padded, blanks.ToArray());
    }

    internal static double[]? ReadNumericData(OpenXmlCompositeElement? valElement)
    {
        if (valElement == null) return null;

        var numLit = valElement.GetFirstChild<C.NumberLiteral>();
        if (numLit != null)
        {
            return numLit.Elements<C.NumericPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => double.TryParse(p.GetFirstChild<C.NumericValue>()?.Text, out var v) ? v : 0)
                .ToArray();
        }

        var numRef = valElement.GetFirstChild<C.NumberReference>();
        var numCache = numRef?.GetFirstChild<C.NumberingCache>();
        if (numCache != null)
        {
            return numCache.Elements<C.NumericPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => double.TryParse(p.GetFirstChild<C.NumericValue>()?.Text, out var v) ? v : 0)
                .ToArray();
        }

        // NumberReference without cache — return empty array (data lives in cells)
        if (numRef != null) return Array.Empty<double>();

        return null;
    }

    /// <summary>
    /// Read the formula string from a NumberReference or StringReference inside a Values/CategoryAxisData element.
    /// Returns null if no reference found.
    /// </summary>
    internal static string? ReadFormulaRef(OpenXmlCompositeElement? element)
    {
        if (element == null) return null;
        var numRef = element.GetFirstChild<C.NumberReference>();
        if (numRef != null) return numRef.GetFirstChild<C.Formula>()?.Text;
        var strRef = element.GetFirstChild<C.StringReference>();
        if (strRef != null) return strRef.GetFirstChild<C.Formula>()?.Text;
        return null;
    }

    internal static string? ReadColorFromFill(Drawing.SolidFill? solidFill)
    {
        if (solidFill == null) return null;
        var rgb = solidFill.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        if (rgb != null) return ParseHelpers.FormatHexColor(rgb);
        var scheme = solidFill.GetFirstChild<Drawing.SchemeColor>()?.Val;
        if (scheme?.HasValue == true) return scheme.InnerText;
        return null;
    }

    /// <summary>
    /// Read any fill child under a chart spPr container (plotArea, chartArea, etc.)
    /// as a Setter-compatible spec string. Recognises a:solidFill (→ "#RRGGBB"),
    /// a:gradFill (→ "c1-c2[:angle]" via ReadGradientSpec), a:noFill (→ "none"),
    /// a:pattFill (→ "pattern:preset[:fg[:bg]]" — same compound form as shape
    /// fills, BuildFillElement reconstructs the full a:pattFill on replay).
    /// Blip fills still round-trip as the literal "blip" hint (no source-part
    /// reconstruction); replay falls back to solid black for that case.
    /// Returns null when the container has no recognised fill child (caller
    /// emits nothing — matches pre-fix behaviour for that case).
    /// </summary>
    internal static string? ReadFillSpec(OpenXmlCompositeElement? spPr)
    {
        if (spPr == null) return null;
        var solid = spPr.GetFirstChild<Drawing.SolidFill>();
        if (solid != null) return ReadColorFromFill(solid);
        var grad = spPr.GetFirstChild<Drawing.GradientFill>();
        if (grad != null) return ReadGradientSpec(grad);
        if (spPr.GetFirstChild<Drawing.NoFill>() != null) return "none";
        var patt = spPr.GetFirstChild<Drawing.PatternFill>();
        if (patt != null) return ReadPatternSpec(patt);
        if (spPr.GetFirstChild<Drawing.BlipFill>() != null) return "blip";
        return null;
    }

    /// <summary>
    /// Read a PatternFill as the dump/replay spec form
    /// "pattern:preset[:fg[:bg]]". Mirrors the shape-side
    /// PowerPointHandler.NodeBuilder pattern emit (which uses "preset:fg:bg"
    /// without the leading hint); the "pattern:" prefix here lets
    /// BuildFillElement disambiguate from solid colors and gradients —
    /// chartFill/plotFill are flat string slots, not a dedicated key like
    /// shape's "pattern" property. Drops alpha; returns "pattern" if the
    /// preset attribute is missing (still recoverable as a coarse hint).
    /// </summary>
    internal static string ReadPatternSpec(Drawing.PatternFill patt)
    {
        var preset = patt.Preset?.InnerText;
        if (string.IsNullOrEmpty(preset)) return "pattern";
        var fgEl = patt.GetFirstChild<Drawing.ForegroundColor>();
        var bgEl = patt.GetFirstChild<Drawing.BackgroundColor>();
        var fg = ReadColorFromColorContainer(fgEl);
        var bg = ReadColorFromColorContainer(bgEl);
        if (fg == null && bg == null) return $"pattern:{preset}";
        if (bg == null) return $"pattern:{preset}:{fg}";
        if (fg == null) fg = "000000";
        return $"pattern:{preset}:{fg}:{bg}";
    }

    private static string? ReadColorFromColorContainer(OpenXmlCompositeElement? el)
    {
        if (el == null) return null;
        var rgb = el.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        if (rgb != null) return ParseHelpers.FormatHexColor(rgb);
        var scheme = el.GetFirstChild<Drawing.SchemeColor>()?.Val;
        if (scheme?.HasValue == true) return scheme.InnerText;
        return null;
    }

    /// <summary>
    /// Read a GradientFill as the dump/replay spec form
    /// "colorA-colorB[-colorC][:angle]". Returns null if no stops can be
    /// resolved. Drops alpha; preserves stop order. Mirrors the input format
    /// accepted by ApplySeriesGradient in the Setter.
    /// </summary>
    internal static string? ReadGradientSpec(Drawing.GradientFill gradFill)
    {
        var stops = gradFill.GetFirstChild<Drawing.GradientStopList>()
            ?.Elements<Drawing.GradientStop>().ToList();
        // R28-B4: a 1-stop gradient is an edge case (Excel/PowerPoint normally
        // require ≥2 stops) but does occur in hand-edited or third-party files.
        // Returning null silently dropped it on dump; instead emit the single
        // color so ApplySeriesGradient (which already tolerates 1-stop input
        // via its duplicate-on-empty fallback) reconstructs an equivalent
        // gradient. Zero stops still cannot round-trip — return null then.
        if (stops == null || stops.Count == 0) return null;
        var parts = new List<string>();
        foreach (var stop in stops)
        {
            var rgb = stop.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            var scheme = stop.GetFirstChild<Drawing.SchemeColor>()?.Val;
            if (rgb != null) parts.Add(rgb);
            else if (scheme?.HasValue == true) parts.Add(scheme.InnerText!);
            else return null;
        }
        var spec = string.Join("-", parts);
        var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
        if (linear?.Angle?.HasValue == true)
            spec += ":" + (linear.Angle.Value / 60000);
        // R63 t-2: <a:lin scaled="0"> round-tripped to scaled="1" because the
        // Builder hard-codes Scaled=true on rebuild. Capture an explicit
        // scaled=false as a ":s0" suffix so BuildFillElement / ApplySeriesGradient
        // can honor the source attribute. Default (omitted or scaled="1")
        // emits nothing — preserves existing dumps. Force the angle slot
        // (":0" when absent) so the suffix doesn't collide with the
        // LastIndexOf(':') angle parser.
        if (linear?.Scaled?.HasValue == true && linear.Scaled.Value == false)
        {
            if (linear.Angle?.HasValue != true) spec += ":0";
            spec += ":s0";
        }
        return spec;
    }

    /// <summary>
    /// Format a dropLines / hiLowLines element as a Setter-spec string.
    /// Empty (no spPr.outline) emits "true"; presence-with-styling emits
    /// "color[:widthPt[:dash]]". Mirrors BuildLineShapeProperties input.
    /// </summary>
    /// <summary>
    /// Format an a:outerShdw element as "COLOR-BLUR-ANGLE-DIST-OPACITY"
    /// — the spec DrawingEffectsHelper.BuildOuterShadow parses. Mirrors
    /// the inverse of that builder so chart series shadow round-trips
    /// through dump→replay.
    /// </summary>
    private static string FormatOuterShadowSpec(Drawing.OuterShadow shadow)
    {
        // Color element (a:srgbClr / a:schemeClr / ...) with optional
        // a:alpha child carrying the opacity in 1000-units.
        var clrEl = shadow.Elements().FirstOrDefault(e =>
            e is Drawing.RgbColorModelHex
                or Drawing.SchemeColor
                or Drawing.PresetColor
                or Drawing.SystemColor
                or Drawing.RgbColorModelPercentage
                or Drawing.HslColor);
        string color;
        if (clrEl is Drawing.RgbColorModelHex rgb && rgb.Val?.HasValue == true)
            color = rgb.Val.Value!;
        else if (clrEl is Drawing.SchemeColor sch && sch.Val?.HasValue == true)
            color = sch.Val.InnerText ?? "000000";
        else
            color = "000000";

        var blurPt = shadow.BlurRadius?.HasValue == true
            ? shadow.BlurRadius.Value / EmuConverter.EmuPerPointF
            : 4.0;
        var distPt = shadow.Distance?.HasValue == true
            ? shadow.Distance.Value / EmuConverter.EmuPerPointF
            : 3.0;
        var angleDeg = shadow.Direction?.HasValue == true
            ? shadow.Direction.Value / 60000.0
            : 45.0;
        var alphaEl = clrEl?.GetFirstChild<Drawing.Alpha>();
        var opacity = alphaEl?.Val?.HasValue == true
            ? alphaEl.Val.Value / 1000.0
            : 40.0;

        string F(double v) => v.ToString("G",
            System.Globalization.CultureInfo.InvariantCulture);
        return $"{color}-{F(blurPt)}-{F(angleDeg)}-{F(distPt)}-{F(opacity)}";
    }

    private static string FormatLineOverlaySpec(OpenXmlElement overlay)
    {
        var spPr = overlay.GetFirstChild<C.ChartShapeProperties>();
        var outline = spPr?.GetFirstChild<Drawing.Outline>();
        if (outline == null) return "true";
        var color = ReadColorFromFill(outline.GetFirstChild<Drawing.SolidFill>());
        if (color == null) return "true";
        // Strip leading '#' for Setter-spec compatibility (BuildChartColorElement
        // accepts both forms; dump output stays consistent with other Setter
        // round-trip keys like updownbars which emit bare hex).
        if (color.StartsWith('#')) color = color.Substring(1);
        var parts = new List<string> { color };
        if (outline.Width?.HasValue == true && outline.Width.Value > 0)
        {
            var widthPt = outline.Width.Value / EmuConverter.EmuPerPointF;
            parts.Add(widthPt.ToString("G", System.Globalization.CultureInfo.InvariantCulture));
        }
        var dash = outline.GetFirstChild<Drawing.PresetDash>()?.Val;
        if (dash?.HasValue == true)
        {
            if (parts.Count == 1) parts.Add("0.5"); // default width slot
            parts.Add(dash.InnerText ?? "");
        }
        return string.Join(":", parts);
    }

    /// <summary>
    /// Format an upDownBars element as "gapWidth[:upColor[:downColor]]".
    /// Mirrors the Setter spec accepted in ChartHelper.Setter.cs case
    /// "updownbars". Empty fill on up/down bars => slot left blank.
    /// </summary>
    private static string FormatUpDownBarsSpec(C.UpDownBars udb)
    {
        var gapWidth = udb.GetFirstChild<C.GapWidth>()?.Val?.Value ?? 150;
        string? upColor = null, downColor = null;
        var upBars = udb.GetFirstChild<C.UpBars>();
        if (upBars != null)
        {
            var fill = upBars.GetFirstChild<C.ChartShapeProperties>()
                ?.GetFirstChild<Drawing.SolidFill>();
            upColor = ReadColorFromFill(fill);
            if (upColor != null && upColor.StartsWith('#')) upColor = upColor.Substring(1);
        }
        var downBars = udb.GetFirstChild<C.DownBars>();
        if (downBars != null)
        {
            var fill = downBars.GetFirstChild<C.ChartShapeProperties>()
                ?.GetFirstChild<Drawing.SolidFill>();
            downColor = ReadColorFromFill(fill);
            if (downColor != null && downColor.StartsWith('#')) downColor = downColor.Substring(1);
        }
        // Only emit trailing slots if non-null. Setter accepts the abbreviated
        // forms ("150", "150:00AA00", "150:00AA00:FF0000").
        if (downColor != null)
            return $"{gapWidth}:{upColor ?? ""}:{downColor}";
        if (upColor != null)
            return $"{gapWidth}:{upColor}";
        return gapWidth.ToString();
    }

    /// <summary>
    /// Build the canonical trendline spec string from a <c:trendline> element.
    /// Embeds order for poly (poly:N) and period for movingAvg (movingAvg:N) so
    /// dump→batch replay round-trips the polynomial degree / window size that
    /// were otherwise silently dropped (the bare type name lost the parameter).
    /// </summary>
    private static string FormatTrendlineSpec(C.Trendline trendline, string typeName)
    {
        if (string.Equals(typeName, "poly", StringComparison.OrdinalIgnoreCase))
        {
            var order = trendline.GetFirstChild<C.PolynomialOrder>()?.Val;
            if (order?.HasValue == true) return $"poly:{order.Value}";
        }
        else if (string.Equals(typeName, "movingAvg", StringComparison.OrdinalIgnoreCase))
        {
            var period = trendline.GetFirstChild<C.Period>()?.Val;
            if (period?.HasValue == true) return $"movingAvg:{period.Value}";
        }
        return typeName;
    }

    /// <summary>
    /// Read gridline detail into separate format keys: {prefix}Color, {prefix}Width, {prefix}Dash.
    /// </summary>
    private static void ReadGridlineDetail(OpenXmlCompositeElement gridlines, DocumentNode node, string prefix)
    {
        var spPr = gridlines.GetFirstChild<C.ChartShapeProperties>();
        var outline = spPr?.GetFirstChild<Drawing.Outline>();
        if (outline == null) return;

        var fill = outline.GetFirstChild<Drawing.SolidFill>();
        var color = ReadColorFromFill(fill);
        if (color != null) node.Format[$"{prefix}Color"] = color;

        if (outline.Width?.HasValue == true)
            node.Format[$"{prefix}Width"] = Math.Round(outline.Width.Value / EmuConverter.EmuPerPointF, 2);

        var dash = outline.GetFirstChild<Drawing.PresetDash>()?.Val;
        if (dash?.HasValue == true)
            node.Format[$"{prefix}Dash"] = dash.InnerText!;
    }

    /// <summary>
    /// Read outline (border) detail into format keys: {prefix}.color, {prefix}.width, {prefix}.dash.
    /// </summary>
    private static void ReadOutlineDetail(Drawing.Outline outline, DocumentNode node, string prefix)
    {
        var fill = outline.GetFirstChild<Drawing.SolidFill>();
        var color = ReadColorFromFill(fill);
        if (color != null) node.Format[$"{prefix}.color"] = color;
        if (outline.Width?.HasValue == true)
            node.Format[$"{prefix}.width"] = Math.Round(outline.Width.Value / EmuConverter.EmuPerPointF, 2);
        var dash = outline.GetFirstChild<Drawing.PresetDash>()?.Val;
        if (dash?.HasValue == true)
            node.Format[$"{prefix}.dash"] = dash.InnerText!;
    }

    /// <summary>
    /// Read font spec from TextProperties: returns "SIZE:COLOR:FONTNAME" format or null.
    /// </summary>
    // An axis title's <a:pPr> is worth round-tripping when its defRPr carries
    // explicit font/size/weight/style/color — i.e. anything that differs from
    // the builder's hard-coded 14pt-bold default — or the paragraph itself sets
    // alignment. Unlike the chart-title check, size (sz) and weight (b) count:
    // axis titles are typically smaller and non-bold, and that is exactly the
    // styling the builder loses.
    private static bool AxisTitlePPrMeaningful(Drawing.ParagraphProperties pPr)
    {
        var defRp = pPr.GetFirstChild<Drawing.DefaultRunProperties>();
        if (defRp != null)
        {
            if (defRp.FontSize?.HasValue == true || defRp.Bold?.HasValue == true
                || defRp.Italic?.HasValue == true)
                return true;
            if (defRp.ChildElements.Any(c => c.LocalName is "solidFill" or "latin" or "ea" or "cs"))
                return true;
            if (defRp.GetAttributes().Any(a => a.LocalName is "u" or "strike"))
                return true;
        }
        return pPr.GetAttributes().Any(a => a.LocalName == "algn");
    }

    private static string? ReadFontSpec(C.TextProperties textProperties)
    {
        var defRp = textProperties.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
        if (defRp == null) return null;

        var parts = new List<string>();
        if (defRp.FontSize?.HasValue == true)
            parts.Add((defRp.FontSize.Value / 100.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture));
        else
            parts.Add("");

        var fill = defRp.GetFirstChild<Drawing.SolidFill>();
        var color = ReadColorFromFill(fill);
        // Canonical: hex colors are emitted with the "#" prefix (project
        // the project conventions). Earlier this stripped "#" via TrimStart, breaking the
        // canonical form for axisFont / legendFont compound readback.
        parts.Add(color ?? "");

        var font = defRp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
        if (font != null)
            parts.Add(font);

        var result = string.Join(":", parts).TrimEnd(':');
        return string.IsNullOrEmpty(result) ? null : result;
    }

    // ==================== Chart Set ====================

    internal static void UpdateSeriesData(C.PlotArea plotArea, List<(string name, double[] values)> newData)
    {
        var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();

        // Update existing series
        for (int i = 0; i < Math.Min(newData.Count, allSer.Count); i++)
        {
            var ser = allSer[i];
            var (sName, sVals) = newData[i];

            var serText = ser.GetFirstChild<C.SeriesText>();
            if (serText != null)
            {
                serText.RemoveAllChildren();
                serText.AppendChild(new C.NumericValue(sName));
            }

            var valEl = ser.GetFirstChild<C.Values>();
            if (valEl != null)
            {
                valEl.RemoveAllChildren();
                var builtVals = BuildValues(sVals);
                foreach (var child in builtVals.ChildElements.ToList())
                    valEl.AppendChild(child.CloneNode(true));
            }
        }

        // Remove excess existing series
        for (int i = newData.Count; i < allSer.Count; i++)
            allSer[i].Remove();

        // Add new series by cloning the last existing one as a template
        if (newData.Count > allSer.Count && allSer.Count > 0)
        {
            var lastSer = allSer[^1];
            var parent = lastSer.Parent!;
            for (int i = allSer.Count; i < newData.Count; i++)
            {
                var (sName, sVals) = newData[i];
                var newSer = (OpenXmlCompositeElement)lastSer.CloneNode(true);

                // Update index and order
                var idx = newSer.GetFirstChild<C.Index>();
                if (idx != null) idx.Val = (uint)i;
                var order = newSer.GetFirstChild<C.Order>();
                if (order != null) order.Val = (uint)i;

                // Update series name
                var serText = newSer.GetFirstChild<C.SeriesText>();
                if (serText != null)
                {
                    serText.RemoveAllChildren();
                    serText.AppendChild(new C.NumericValue(sName));
                }

                // Update values
                var valEl = newSer.GetFirstChild<C.Values>();
                if (valEl != null)
                {
                    valEl.RemoveAllChildren();
                    var builtVals = BuildValues(sVals);
                    foreach (var child in builtVals.ChildElements.ToList())
                        valEl.AppendChild(child.CloneNode(true));
                }

                // Remove cloned color so the new series gets a distinct auto-color
                var spPr = newSer.GetFirstChild<C.ChartShapeProperties>();
                if (spPr != null) spPr.Remove();

                parent.AppendChild(newSer);
            }
        }
    }
}
