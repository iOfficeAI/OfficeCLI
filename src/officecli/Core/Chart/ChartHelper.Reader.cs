// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

internal static partial class ChartHelper
{
    // ==================== Chart Readback ====================

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
        var titleText = titleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (titleText == null && titleEl != null)
        {
            // BuildChartTitle routes single-cell-reference values (e.g. "Q1",
            // "Sheet1!A1") through a <c:strRef><c:f>...</c:f></c:strRef> path
            // instead of <a:t> literal text. Surface the formula so a get→set
            // round-trip preserves the reference and the schema-declared
            // 'title' get readback isn't silently empty.
            var strRefFormula = titleEl.Descendants<C.Formula>().FirstOrDefault()?.Text;
            if (!string.IsNullOrEmpty(strRefFormula)) titleText = strRefFormula;
        }
        if (titleText != null) node.Format["title"] = titleText;

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
            }
        }

        var legend = chart.GetFirstChild<C.Legend>();
        if (legend != null)
        {
            var posRaw = legend.GetFirstChild<C.LegendPosition>()?.Val?.HasValue == true
                ? legend.GetFirstChild<C.LegendPosition>()!.Val!.InnerText : "b";
            node.Format["legend"] = posRaw switch
            {
                "b" => "bottom",
                "t" => "top",
                "l" => "left",
                "r" => "right",
                "tr" => "topRight",
                _ => posRaw
            };
        }
        else
        {
            // Builder defaults to legend=bottom when prop absent; emit explicit
            // "none" so dump→replay round-trip preserves the no-legend state.
            node.Format["legend"] = "none";
        }

        var dataLabels = plotArea.Descendants<C.DataLabels>().FirstOrDefault();
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
                node.Format["labelPos"] = dlPos.InnerText;
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
        var plotSpPr = plotArea.GetFirstChild<C.ShapeProperties>();
        var plotFill = plotSpPr?.GetFirstChild<Drawing.SolidFill>();
        if (plotFill != null)
        {
            var pColor = ReadColorFromFill(plotFill);
            if (pColor != null) node.Format["plotFill"] = pColor;
        }

        // Chart area fill (ChartSpace > spPr, NOT PlotArea)
        // Note: The SDK serializes ChartShapeProperties but deserializes it as C.ShapeProperties
        // after round-trip. Check both types, plus in-memory ChartShapeProperties.
        {
            Drawing.SolidFill? chartAreaFill = null;
            var csSpPr = chart.Parent?.GetFirstChild<C.ShapeProperties>();
            if (csSpPr != null)
                chartAreaFill = csSpPr.GetFirstChild<Drawing.SolidFill>();
            if (chartAreaFill == null)
            {
                var csCSpPr = chart.Parent?.GetFirstChild<C.ChartShapeProperties>();
                if (csCSpPr != null)
                    chartAreaFill = csCSpPr.GetFirstChild<Drawing.SolidFill>();
            }
            if (chartAreaFill != null)
            {
                var cColor = ReadColorFromFill(chartAreaFill);
                if (cColor != null) node.Format["chartFill"] = cColor;
            }
        }

        // Gridlines: "true" for presence, detail in gridlineColor/gridlineWidth/gridlineDash
        var valAxisForGrid = plotArea.GetFirstChild<C.ValueAxis>();
        var majorGL = valAxisForGrid?.GetFirstChild<C.MajorGridlines>();
        if (majorGL != null)
        {
            node.Format["gridlines"] = "true";
            ReadGridlineDetail(majorGL, node, "gridline");
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
        }

        // GapWidth / Overlap from bar/column chart
        var barChart = plotArea.GetFirstChild<C.BarChart>();
        var gapWidthEl = barChart?.GetFirstChild<C.GapWidth>();
        if (gapWidthEl?.Val?.HasValue == true) node.Format["gapwidth"] = gapWidthEl.Val.Value.ToString();
        var overlapEl = barChart?.GetFirstChild<C.Overlap>();
        if (overlapEl?.Val?.HasValue == true) node.Format["overlap"] = overlapEl.Val.Value.ToString();

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

        // Secondary axis — emit the 1-based series indices bound to the
        // secondary axis so dump→replay round-trips. The Setter expects
        // "1,3" form (series indices); emitting bare "true" silently failed
        // parsing on replay because every comma-split token tried as int
        // produced [-1] then was filtered out.
        var valAxes = plotArea.Elements<C.ValueAxis>().ToList();
        if (valAxes.Count > 1)
        {
            // Map AxisId -> rank by document order; rank 0 = primary, 1 = secondary.
            var axisRank = new Dictionary<uint, int>();
            for (int ai = 0; ai < valAxes.Count; ai++)
            {
                var axId = valAxes[ai].GetFirstChild<C.AxisId>()?.Val?.Value;
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
                        .Select(a => a.Val?.Value)
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
        var valAxisFirst = plotArea.GetFirstChild<C.ValueAxis>();
        var valAxisTxPrRot = valAxisFirst?.GetFirstChild<C.TextProperties>();
        var valAxisBodyPr = valAxisTxPrRot?.GetFirstChild<Drawing.BodyProperties>();
        if (valAxisBodyPr?.Rotation?.HasValue == true)
        {
            var deg = valAxisBodyPr.Rotation.Value / 60000.0;
            node.Format["yaxis.labelRotation"] = deg.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        }

        // Axis titles
        var valAxis = plotArea.GetFirstChild<C.ValueAxis>();
        var valAxisTitle = valAxis?.GetFirstChild<C.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (valAxisTitle != null) node.Format["axisTitle"] = valAxisTitle;

        var catAxis = plotArea.GetFirstChild<C.CategoryAxis>();
        var catAxisTitle = catAxis?.GetFirstChild<C.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (catAxisTitle != null) node.Format["catTitle"] = catAxisTitle;

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
        }

        // Chart-type-specific
        var pieChart = plotArea.GetFirstChild<C.PieChart>();
        var firstSliceAngle = pieChart?.GetFirstChild<C.FirstSliceAngle>()?.Val?.Value;
        if (firstSliceAngle != null && firstSliceAngle != 0) node.Format["firstSliceAngle"] = firstSliceAngle;

        var doughnutChart = plotArea.GetFirstChild<C.DoughnutChart>();
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
                    node.Format["labelFont.size"] = (dlDefRp.FontSize.Value / 100).ToString();
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

                // Cell reference formulas (for series with NumberReference/StringReference)
                if (serEl != null)
                {
                    var valRef = ReadFormulaRef(serEl.GetFirstChild<C.Values>());
                    if (valRef != null) seriesNode.Format["valuesRef"] = valRef;
                    var catRef = ReadFormulaRef(serEl.GetFirstChild<C.CategoryAxisData>());
                    if (catRef != null) seriesNode.Format["categoriesRef"] = catRef;
                    var nameRefF = serEl.GetFirstChild<C.SeriesText>()
                        ?.GetFirstChild<C.StringReference>()
                        ?.GetFirstChild<C.Formula>()?.Text;
                    if (!string.IsNullOrEmpty(nameRefF)) seriesNode.Format["nameRef"] = nameRefF;
                }

                var serSpPr = serEl?.GetFirstChild<C.ChartShapeProperties>();
                // NoFill round-trip: when ApplySeriesColor wrote <a:noFill/>
                // (color=none), Reader previously skipped emit and dump→replay
                // reverted to the default auto color. Surface "none" so the
                // setter side re-applies NoFill.
                if (serSpPr?.GetFirstChild<Drawing.NoFill>() != null)
                    seriesNode.Format["color"] = "none";
                var serColor = serSpPr?.GetFirstChild<Drawing.SolidFill>();
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
                    seriesNode.Format["lineWidth"] = Math.Round(outline.Width.Value / 12700.0, 2);
                // Line dash
                var prstDash = outline?.GetFirstChild<Drawing.PresetDash>();
                if (prstDash?.Val?.HasValue == true)
                    seriesNode.Format["lineDash"] = prstDash.Val.InnerText;
                // Outline color
                var outlineFill = outline?.GetFirstChild<Drawing.SolidFill>();
                if (outlineFill != null)
                {
                    var outColor = ReadColorFromFill(outlineFill);
                    if (outColor != null) seriesNode.Format["outlineColor"] = outColor;
                }
                // Shadow (from EffectList)
                var effectList = serSpPr?.GetFirstChild<Drawing.EffectList>();
                var outerShadow = effectList?.GetFirstChild<Drawing.OuterShadow>();
                if (outerShadow != null) seriesNode.Format["shadow"] = "true";
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
                            "stdDev" => "stddev",
                            "stdErr" => "stderr",
                            _ => errValType.InnerText
                        };
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
                        seriesNode.Format["errbars"] = mag != null
                            ? $"{typeName}:{mag}"
                            : typeName;
                    }
                }
                // InvertIfNegative
                var inv = serEl?.GetFirstChild<C.InvertIfNegative>()?.Val;
                if (inv?.HasValue == true && inv.Value) seriesNode.Format["invertIfNeg"] = "true";
                // Explosion (pie)
                var explosion = serEl?.GetFirstChild<C.Explosion>()?.Val?.Value;
                if (explosion != null && explosion > 0) seriesNode.Format["explosion"] = explosion;
                // Data point colors
                if (serEl != null)
                {
                    foreach (var dPt in serEl.Elements<C.DataPoint>())
                    {
                        var ptIdx = dPt.Index?.Val?.Value;
                        if (ptIdx == null) continue;
                        var ptSpPr = dPt.GetFirstChild<C.ChartShapeProperties>();
                        if (ptSpPr?.GetFirstChild<Drawing.NoFill>() != null)
                            seriesNode.Format[$"point{ptIdx.Value + 1}.color"] = "none";
                        var ptFill = ptSpPr?.GetFirstChild<Drawing.SolidFill>();
                        if (ptFill != null)
                        {
                            var ptColor = ReadColorFromFill(ptFill);
                            if (ptColor != null) seriesNode.Format[$"point{ptIdx.Value + 1}.color"] = ptColor;
                        }
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
    internal static List<(string Name, double Value, string Color, double WidthPt, string Dash)> ReadReferenceLines(C.PlotArea plotArea)
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
                var widthPt = widthEmu / 12700.0;

                // Color: solidFill srgbClr val
                var color = "FF0000";
                var srgb = outline?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                if (!string.IsNullOrEmpty(srgb)) color = srgb;

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
        if (catData == null) return null;

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

    internal static List<(string name, double[] values)> ReadAllSeries(C.PlotArea plotArea)
    {
        var result = new List<(string name, double[] values)>();

        foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent != null &&
                (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart"))))
        {
            var serText = ser.GetFirstChild<C.SeriesText>();
            // c:tx may carry <c:strRef> (cached cell value) or <c:v> (literal).
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
                name = serText?.Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? "?";
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
        return spec;
    }

    /// <summary>
    /// Format a dropLines / hiLowLines element as a Setter-spec string.
    /// Empty (no spPr.outline) emits "true"; presence-with-styling emits
    /// "color[:widthPt[:dash]]". Mirrors BuildLineShapeProperties input.
    /// </summary>
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
            var widthPt = outline.Width.Value / 12700.0;
            parts.Add(widthPt.ToString("G", System.Globalization.CultureInfo.InvariantCulture));
        }
        var dash = outline.GetFirstChild<Drawing.PresetDash>()?.Val;
        if (dash?.HasValue == true)
        {
            if (parts.Count == 1) parts.Add("0.5"); // default width slot
            parts.Add(dash.InnerText);
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
            node.Format[$"{prefix}Width"] = Math.Round(outline.Width.Value / 12700.0, 2);

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
            node.Format[$"{prefix}.width"] = Math.Round(outline.Width.Value / 12700.0, 2);
        var dash = outline.GetFirstChild<Drawing.PresetDash>()?.Val;
        if (dash?.HasValue == true)
            node.Format[$"{prefix}.dash"] = dash.InnerText!;
    }

    /// <summary>
    /// Read font spec from TextProperties: returns "SIZE:COLOR:FONTNAME" format or null.
    /// </summary>
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
        parts.Add(color?.TrimStart('#') ?? "");

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
