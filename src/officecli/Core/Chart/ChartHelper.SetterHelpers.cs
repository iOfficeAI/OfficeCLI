// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

/// <summary>
/// Additional helper methods for ChartSetter — split out to keep file sizes manageable.
/// Covers: tick marks, trendlines, error bars, borders, data point styling.
/// </summary>
internal static partial class ChartHelper
{
    // ==================== Legend Position ====================

    /// <summary>
    /// Parse a user-supplied legend position string into the OOXML enum.
    /// Throws ArgumentException on unknown tokens — historically these
    /// silently fell through to "bottom", producing a contradictory
    /// "Updated: legend=hidden" success message while the file actually
    /// carried legend=bottom (R34-1). Caller should already have handled
    /// "none" / "false" (legend removal) before reaching here.
    /// </summary>
    internal static C.LegendPositionValues ParseLegendPosition(string value)
    {
        // CONSISTENCY(legend-separator-normalize): accept dash AND underscore
        // as separators (`top-right`, `top_right`, `TOP_RIGHT`) by stripping
        // both before comparison. Without this, `TOP_RIGHT` threw while
        // `top-right` succeeded — punctuation variants should be uniform.
        var norm = (value ?? string.Empty).ToLowerInvariant().Replace("-", "").Replace("_", "");
        return norm switch
        {
            "top" or "t" => C.LegendPositionValues.Top,
            "bottom" or "b" => C.LegendPositionValues.Bottom,
            "left" or "l" => C.LegendPositionValues.Left,
            "right" or "r" => C.LegendPositionValues.Right,
            "topright" or "tr" => C.LegendPositionValues.TopRight,
            _ => throw new ArgumentException(
                $"Invalid legend position '{value}'. " +
                "Valid: none, top, bottom, left, right, topRight " +
                "(or use 'none'/'false' to hide the legend)."),
        };
    }

    // ==================== Tick Mark Helpers ====================

    internal static C.TickMarkValues ParseTickMark(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "none" or "false" => C.TickMarkValues.None,
            "in" or "inside" => C.TickMarkValues.Inside,
            "out" or "outside" => C.TickMarkValues.Outside,
            "cross" or "both" => C.TickMarkValues.Cross,
            _ => throw new ArgumentException(
                $"Invalid tick mark value '{value}'. Valid values: none, in, out, cross.")
        };
    }

    // ==================== Trendline Helpers ====================

    internal static C.Trendline BuildTrendline(string spec)
    {
        // Format: "type" or "type:order" or "type:forward:backward"
        // e.g. "linear", "poly:3", "exp:2:1", "movingAvg:3"
        var parts = spec.Split(':');
        var typeStr = parts[0].Trim().ToLowerInvariant();

        var trendline = new C.Trendline();

        var trendType = typeStr switch
        {
            "exp" or "exponential" => C.TrendlineValues.Exponential,
            "log" or "logarithmic" => C.TrendlineValues.Logarithmic,
            "poly" or "polynomial" => C.TrendlineValues.Polynomial,
            "power" => C.TrendlineValues.Power,
            "movingavg" or "moving" or "movingaverage" => C.TrendlineValues.MovingAverage,
            _ => C.TrendlineValues.Linear
        };
        trendline.AppendChild(new C.TrendlineType { Val = trendType });

        // Polynomial order or moving average period
        if (parts.Length > 1 && int.TryParse(parts[1], out var order))
        {
            if (trendType == C.TrendlineValues.Polynomial)
                trendline.AppendChild(new C.PolynomialOrder { Val = (byte)Math.Clamp(order, 2, 6) });
            else if (trendType == C.TrendlineValues.MovingAverage)
                trendline.AppendChild(new C.Period { Val = (uint)order });
            else
            {
                // Treat as forward extrapolation periods
                trendline.AppendChild(new C.Forward { Val = order });
            }
        }

        // Backward extrapolation
        if (parts.Length > 2 && double.TryParse(parts[2],
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var backward))
        {
            trendline.AppendChild(new C.Backward { Val = backward });
        }

        return trendline;
    }

    internal static void ApplyTrendlineOptions(C.Trendline trendline, string optionKey, string value)
    {
        switch (optionKey)
        {
            case "name" or "label":
                trendline.RemoveAllChildren<C.TrendlineName>();
                trendline.PrependChild(new C.TrendlineName { Text = value });
                // Also emit a <c:trendlineLbl> with rich-text so Excel actually
                // paints the label next to the trendline (a <c:name> alone is
                // used by older tooling as a legend-entry override).
                trendline.RemoveAllChildren<C.TrendlineLabel>();
                var tlLbl = new C.TrendlineLabel(
                    new C.Layout(),
                    new C.ChartText(
                        new C.RichText(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph(
                                new Drawing.Run(
                                    new Drawing.RunProperties { Language = "en-US" },
                                    new Drawing.Text(value))))));
                // Schema order under CT_Trendline: name, trendlineLbl, trendlineType, ...
                var trendlineType = trendline.GetFirstChild<C.TrendlineType>();
                if (trendlineType != null)
                    trendline.InsertBefore(tlLbl, trendlineType);
                else
                    trendline.AppendChild(tlLbl);
                break;
            case "forward" or "forecastforward":
                trendline.RemoveAllChildren<C.Forward>();
                trendline.AppendChild(new C.Forward { Val = ParseHelpers.SafeParseDouble(value, "trendline.forward") });
                break;
            case "backward" or "forecastbackward":
                trendline.RemoveAllChildren<C.Backward>();
                trendline.AppendChild(new C.Backward { Val = ParseHelpers.SafeParseDouble(value, "trendline.backward") });
                break;
            case "order":
                trendline.RemoveAllChildren<C.PolynomialOrder>();
                trendline.AppendChild(new C.PolynomialOrder { Val = (byte)Math.Clamp(ParseHelpers.SafeParseInt(value, "trendline.order"), 2, 6) });
                break;
            case "period":
                trendline.RemoveAllChildren<C.Period>();
                trendline.AppendChild(new C.Period { Val = (uint)Math.Max(2, ParseHelpers.SafeParseInt(value, "trendline.period")) });
                break;
            case "intercept":
                trendline.RemoveAllChildren<C.Intercept>();
                trendline.AppendChild(new C.Intercept { Val = ParseHelpers.SafeParseDouble(value, "trendline.intercept") });
                break;
            case "disprsqr" or "rsquared" or "r2" or "displayrsquared":
                trendline.RemoveAllChildren<C.DisplayRSquaredValue>();
                trendline.AppendChild(new C.DisplayRSquaredValue { Val = ParseHelpers.IsTruthy(value) });
                break;
            case "dispeq" or "equation" or "displayequation":
                trendline.RemoveAllChildren<C.DisplayEquation>();
                trendline.AppendChild(new C.DisplayEquation { Val = ParseHelpers.IsTruthy(value) });
                break;
        }
    }

    // ==================== Error Bars Helpers ====================

    /// <summary>
    /// Check if the parent chart type supports errBars on its series (CT_*Ser).
    /// OOXML allows errBars in: barChart, bar3DChart, scatterChart, areaChart,
    /// area3DChart, bubbleChart.  Not allowed in: lineChart, line3DChart,
    /// pieChart, pie3DChart, doughnutChart, radarChart, stockChart.
    /// </summary>
    internal static bool SeriesSupportsErrorBars(OpenXmlElement ser)
    {
        var parentName = ser.Parent?.LocalName ?? "";
        return parentName is "barChart" or "bar3DChart"
            or "scatterChart"
            or "areaChart" or "area3DChart"
            or "bubbleChart";
    }

    internal static C.ErrorBars BuildErrorBars(string spec)
    {
        // Format: "type" or "type:value" e.g. "fixed:5", "percent:10", "stddev", "stderr"
        var parts = spec.Split(':');
        var typeStr = parts[0].Trim().ToLowerInvariant();

        var errBars = new C.ErrorBars();
        errBars.AppendChild(new C.ErrorDirection { Val = C.ErrorBarDirectionValues.Y });
        errBars.AppendChild(new C.ErrorBarType { Val = C.ErrorBarValues.Both });

        var errValType = typeStr switch
        {
            "fixed" or "fixedvalue" => C.ErrorValues.FixedValue,
            "percent" or "pct" => C.ErrorValues.Percentage,
            "stddev" or "standarddeviation" => C.ErrorValues.StandardDeviation,
            "stderr" or "standarderror" => C.ErrorValues.StandardError,
            _ => C.ErrorValues.FixedValue
        };
        errBars.AppendChild(new C.ErrorBarValueType { Val = errValType });

        if (parts.Length > 1 && double.TryParse(parts[1],
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var errVal))
        {
            var numLit = new C.NumberLiteral(
                new C.FormatCode("General"),
                new C.PointCount { Val = 1 },
                new C.NumericPoint(new C.NumericValue(errVal.ToString("G"))) { Index = 0 });
            errBars.AppendChild(new C.Plus(numLit));
            errBars.AppendChild(new C.Minus(numLit.CloneNode(true)));
        }

        return errBars;
    }

    // ==================== Border / Outline Helpers ====================

    internal static Drawing.Outline BuildOutlineElement(string spec)
    {
        // Format: "color" or "color:width" or "color:width:dash"
        // e.g. "000000", "333333:1.5", "666666:1:dash"
        var parts = spec.Split(':');
        var color = parts[0].Trim();
        var widthPt = parts.Length > 1 && double.TryParse(parts[1],
            System.Globalization.CultureInfo.InvariantCulture, out var w) ? w : 0.75;
        var dash = parts.Length > 2 ? parts[2].Trim() : null;

        var outline = new Drawing.Outline { Width = (int)(widthPt * 12700) };
        var sf = new Drawing.SolidFill();
        sf.AppendChild(BuildChartColorElement(color));
        outline.AppendChild(sf);

        if (!string.IsNullOrEmpty(dash))
            outline.AppendChild(new Drawing.PresetDash { Val = ParseDashStyle(dash) });

        return outline;
    }

    // ==================== Per-Series Data Point Helpers ====================

    internal static void ApplyDataPointColor(OpenXmlCompositeElement series, int pointIndex, string color)
    {
        // Find or create c:dPt with the matching index (0-based)
        var dPts = series.Elements<C.DataPoint>().ToList();
        var dPt = dPts.FirstOrDefault(dp => dp.Index?.Val?.Value == (uint)pointIndex);
        if (dPt == null)
        {
            dPt = new C.DataPoint();
            dPt.AppendChild(new C.Index { Val = (uint)pointIndex });
            // Insert before c:dLbls, c:trendline, c:errBars, c:cat, c:val etc.
            var insertBefore = series.GetFirstChild<C.DataLabels>() as OpenXmlElement
                ?? series.GetFirstChild<C.Trendline>() as OpenXmlElement
                ?? series.GetFirstChild<C.ErrorBars>() as OpenXmlElement
                ?? series.GetFirstChild<C.CategoryAxisData>() as OpenXmlElement
                ?? series.GetFirstChild<C.Values>();
            if (insertBefore != null)
                series.InsertBefore(dPt, insertBefore);
            else
                series.AppendChild(dPt);
        }

        var spPr = dPt.GetFirstChild<C.ChartShapeProperties>();
        if (spPr == null) { spPr = new C.ChartShapeProperties(); dPt.AppendChild(spPr); }
        spPr.RemoveAllChildren<Drawing.SolidFill>();
        var fill = new Drawing.SolidFill();
        fill.AppendChild(BuildChartColorElement(color));
        spPr.PrependChild(fill);
    }

    internal static void ApplyDataPointExplosion(OpenXmlCompositeElement series, int pointIndex, uint explosion)
    {
        var dPts = series.Elements<C.DataPoint>().ToList();
        var dPt = dPts.FirstOrDefault(dp => dp.Index?.Val?.Value == (uint)pointIndex);
        if (dPt == null)
        {
            dPt = new C.DataPoint();
            dPt.AppendChild(new C.Index { Val = (uint)pointIndex });
            var insertBefore = series.GetFirstChild<C.DataLabels>() as OpenXmlElement
                ?? series.GetFirstChild<C.CategoryAxisData>() as OpenXmlElement
                ?? series.GetFirstChild<C.Values>();
            if (insertBefore != null) series.InsertBefore(dPt, insertBefore);
            else series.AppendChild(dPt);
        }
        dPt.RemoveAllChildren<C.Explosion>();
        if (explosion > 0)
            dPt.AppendChild(new C.Explosion { Val = explosion });
    }

    // ==================== Axis Line Styling ====================

    /// <summary>
    /// Apply outline (line style) to an axis element's own ShapeProperties.
    /// Format: "color" or "color:width" or "color:width:dash" or "none"
    /// </summary>
    internal static void ApplyAxisLine(OpenXmlCompositeElement axis, string value)
    {
        var spPr = axis.GetFirstChild<C.ChartShapeProperties>();
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            if (spPr != null)
            {
                spPr.RemoveAllChildren<Drawing.Outline>();
                var outline = new Drawing.Outline();
                outline.AppendChild(new Drawing.NoFill());
                spPr.AppendChild(outline);
            }
            return;
        }

        if (spPr == null)
        {
            spPr = new C.ChartShapeProperties();
            // Insert after TickLabelPosition or at end
            var tlPos = axis.GetFirstChild<C.TickLabelPosition>();
            if (tlPos != null) tlPos.InsertAfterSelf(spPr);
            else axis.AppendChild(spPr);
        }
        spPr.RemoveAllChildren<Drawing.Outline>();
        spPr.AppendChild(BuildOutlineElement(value));
    }

    // ==================== Dotted Key Parsers ====================

    /// <summary>
    /// Parse keys like "series1.smooth", "series2.trendline", "series1.point2.color".
    /// Returns (seriesIndex, propertyPath) e.g. (1, "smooth") or (1, "point2.color").
    /// </summary>
    internal static bool TryParseSeriesDottedKey(string key, out int seriesIndex, out string property)
    {
        seriesIndex = 0;
        property = "";
        var lower = key.ToLowerInvariant();
        if (!lower.StartsWith("series")) return false;
        var rest = lower["series".Length..]; // e.g. "1.smooth"
        var dotIdx = rest.IndexOf('.');
        if (dotIdx <= 0) return false;
        if (!int.TryParse(rest[..dotIdx], out seriesIndex) || seriesIndex < 1) return false;
        property = rest[(dotIdx + 1)..];
        return !string.IsNullOrEmpty(property);
    }

    /// <summary>
    /// Handle per-series dotted properties: smooth, trendline, trendline.*, marker, markerSize,
    /// point{M}.color, point{M}.explosion, invertIfNeg, errBars, color.
    /// Returns true if the property was recognized and handled; false otherwise so the
    /// caller can surface it as "unsupported" rather than silently accepting it.
    /// </summary>
    internal static bool HandleSeriesDottedProperty(OpenXmlCompositeElement ser, string prop, string value)
    {
        switch (prop)
        {
            case "smooth":
                // smooth only valid on line/scatter series (CT_LineSer, CT_ScatterSer)
                if (ser.Parent is C.LineChart or C.ScatterChart)
                {
                    ser.RemoveAllChildren<C.Smooth>();
                    InsertSeriesChildInOrder(ser, new C.Smooth { Val = ParseHelpers.IsTruthy(value) });
                }
                return true;

            case "trendline":
                // CL20: `Set trendline=X` APPENDS a trendline (Excel allows
                // multiple trendlines per series). Pass `none` to clear.
                // If the requested trendline type already exists on the
                // series, replace it in place so repeated identical sets
                // stay idempotent; otherwise append a new one.
                if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    ser.RemoveAllChildren<C.Trendline>();
                }
                else
                {
                    var newTl = BuildTrendline(value);
                    var newType = newTl.GetFirstChild<C.TrendlineType>()?.Val?.Value;
                    var dupeTl = ser.Elements<C.Trendline>()
                        .FirstOrDefault(t => t.GetFirstChild<C.TrendlineType>()?.Val?.Value == newType);
                    if (dupeTl != null)
                    {
                        dupeTl.InsertAfterSelf(newTl);
                        dupeTl.Remove();
                    }
                    else
                    {
                        InsertSeriesChildInOrder(ser, newTl);
                    }
                }
                return true;

            case "marker":
                ApplySeriesMarker(ser, value);
                return true;

            case "markersize":
            case "marker.size":
            {
                var marker = ser.GetFirstChild<C.Marker>();
                if (marker == null)
                {
                    marker = new C.Marker();
                    var insertBefore = (OpenXmlElement?)ser.Elements().FirstOrDefault(e =>
                        e.LocalName is "xVal" or "yVal" or "cat" or "val" or "bubbleSize"
                            or "smooth" or "extLst")
                        ?? ser.Elements().FirstOrDefault(e => e.LocalName == "trendline");
                    if (insertBefore != null) ser.InsertBefore(marker, insertBefore);
                    else ser.AppendChild(marker);
                }
                marker.RemoveAllChildren<C.Size>();
                marker.AppendChild(new C.Size { Val = ParseHelpers.SafeParseByte(value, "series.markerSize") });
                return true;
            }

            case "marker.style":
            {
                // CONSISTENCY(marker-dotted): mirror "marker=circle" but accept the
                // dotted alternative seriesN.marker.style=circle. Preserve any
                // existing c:size so users can set style and size independently.
                var existing = ser.GetFirstChild<C.Marker>();
                var existingSize = existing?.GetFirstChild<C.Size>()?.Val?.Value;
                ApplySeriesMarker(ser, value);
                if (existingSize.HasValue)
                {
                    var newMarker = ser.GetFirstChild<C.Marker>();
                    if (newMarker != null && newMarker.GetFirstChild<C.Size>() == null)
                    {
                        var sym = newMarker.GetFirstChild<C.Symbol>();
                        var sz = new C.Size { Val = existingSize.Value };
                        if (sym != null) sym.InsertAfterSelf(sz);
                        else newMarker.AppendChild(sz);
                    }
                }
                return true;
            }

            case "color":
                ApplySeriesColor(ser, value);
                return true;

            case "name":
            {
                var serText = ser.GetFirstChild<C.SeriesText>();
                if (serText != null)
                {
                    // If the value looks like a cell reference, rewrite c:tx as a
                    // c:strRef so Excel resolves it to the cell's value (matches
                    // Add-path behavior for series{N}.name=Sheet1!A1).
                    if (IsCellReference(value))
                    {
                        RewriteSeriesTextAsRef(ser, NormalizeCellReference(value), cachedValue: null);
                    }
                    else
                    {
                        serText.RemoveAllChildren();
                        serText.AppendChild(new C.NumericValue(value));
                    }
                }
                return true;
            }

            case "values":
            {
                var valEl = ser.GetFirstChild<C.Values>();
                if (valEl != null)
                {
                    valEl.RemoveAllChildren();
                    if (value.Contains('!'))
                    {
                        // Cell reference: e.g. Sheet1!B2:B4
                        var builtVals = BuildValuesRef(value);
                        foreach (var child in builtVals.ChildElements.ToList())
                            valEl.AppendChild(child.CloneNode(true));
                    }
                    else
                    {
                        var nums = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                            .Select(s => double.TryParse(s, System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : 0.0)
                            .ToArray();
                        var builtVals = BuildValues(nums);
                        foreach (var child in builtVals.ChildElements.ToList())
                            valEl.AppendChild(child.CloneNode(true));
                    }
                }
                return true;
            }

            case "invertifneg" or "invertifnegative":
                ser.RemoveAllChildren<C.InvertIfNegative>();
                ser.AppendChild(new C.InvertIfNegative { Val = ParseHelpers.IsTruthy(value) });
                return true;

            case "errbars" or "errorbars":
                ser.RemoveAllChildren<C.ErrorBars>();
                if (!value.Equals("none", StringComparison.OrdinalIgnoreCase)
                    && SeriesSupportsErrorBars(ser))
                    InsertSeriesChildInOrder(ser, BuildErrorBars(value));
                return true;

            case "explosion" or "explode":
                ser.RemoveAllChildren<C.Explosion>();
                if (uint.TryParse(value, out var expVal) && expVal > 0)
                    ser.AppendChild(new C.Explosion { Val = expVal });
                return true;

            case "linewidth":
                ApplySeriesLineWidth(ser, (int)(ParseHelpers.SafeParseDouble(value, "series.lineWidth") * 12700));
                return true;

            case "linedash" or "dash":
                ApplySeriesLineDash(ser, value);
                return true;

            case "shadow":
            {
                var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
                if (spPr == null) { spPr = new C.ChartShapeProperties(); ser.AppendChild(spPr); }
                var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? new Drawing.EffectList();
                if (effectList.Parent == null)
                    InsertEffectListInChartSpPr(spPr, effectList);
                effectList.RemoveAllChildren<Drawing.OuterShadow>();
                if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    effectList.AppendChild(DrawingEffectsHelper.BuildOuterShadow(value, BuildChartColorElement));
                return true;
            }

            case "outline":
            {
                var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
                if (spPr == null) { spPr = new C.ChartShapeProperties(); ser.AppendChild(spPr); }
                spPr.RemoveAllChildren<Drawing.Outline>();
                if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    var outlineEl = BuildOutlineElement(value);
                    var effLst = spPr.GetFirstChild<Drawing.EffectList>();
                    if (effLst != null) spPr.InsertBefore(outlineEl, effLst);
                    else spPr.AppendChild(outlineEl);
                }
                return true;
            }

            case "gradient" or "gradientfill":
                ApplySeriesGradient(ser, value);
                return true;

            case "alpha" or "transparency":
            {
                var alphaPercent = ParseHelpers.SafeParseDouble(value, "series.alpha");
                if (prop == "transparency") alphaPercent = 100.0 - alphaPercent;
                ApplySeriesAlpha(ser, (int)(alphaPercent * 1000));
                return true;
            }

            // R26-2: `series{N}.displayEquation` / `series{N}.displayRSquared`
            // are convenience aliases that target the series' existing
            // trendline (equivalent to `series{N}.trendline.displayEquation`).
            // Mirrors the chart-level `trendline.displayequation` fan-out.
            case "displayequation" or "equation" or "dispeq":
            case "displayrsquared" or "rsquared" or "r2" or "disprsqr":
            {
                var tl = ser.GetFirstChild<C.Trendline>();
                if (tl == null) return false;
                ApplyTrendlineOptions(tl, prop, value);
                return true;
            }

            default:
                // Trendline sub-properties: series{N}.trendline.name, .forward, .backward, etc.
                // NOTE: this is an inner dispatch — if the sub-property is not one of
                // ApplyTrendlineOptions' known cases it is silently ignored. See report:
                // same silent-accept risk exists for trendline.* and point{M}.* sub-keys.
                if (prop.StartsWith("trendline."))
                {
                    var tl = ser.GetFirstChild<C.Trendline>();
                    if (tl != null)
                        ApplyTrendlineOptions(tl, prop["trendline.".Length..], value);
                    return true;
                }
                // Per-point properties: series{N}.point{M}.color, series{N}.point{M}.explosion
                if (prop.StartsWith("point") && TryParsePointKey(prop, out var ptIdx, out var ptProp))
                {
                    switch (ptProp)
                    {
                        case "color":
                            ApplyDataPointColor(ser, ptIdx - 1, value);
                            return true;
                        case "explosion" or "explode":
                            ApplyDataPointExplosion(ser, ptIdx - 1,
                                uint.TryParse(value, out var pe) ? pe : 0u);
                            return true;
                        default:
                            // Unknown point sub-property — surface as unsupported.
                            return false;
                    }
                }
                // Genuinely unknown series sub-property (e.g. chartType, axisGroup) —
                // surface via `unsupported` so callers see "Set lied" errors instead
                // of a bogus "Updated" message.
                return false;
        }
    }

    private static bool TryParsePointKey(string prop, out int pointIndex, out string pointProp)
    {
        // Parse "point2.color" → (2, "color")
        pointIndex = 0;
        pointProp = "";
        if (!prop.StartsWith("point")) return false;
        var rest = prop["point".Length..];
        var dotIdx = rest.IndexOf('.');
        if (dotIdx <= 0) return false;
        if (!int.TryParse(rest[..dotIdx], out pointIndex) || pointIndex < 1) return false;
        pointProp = rest[(dotIdx + 1)..];
        return !string.IsNullOrEmpty(pointProp);
    }

    /// <summary>
    /// Parse keys like "dataLabel1.delete", "dataLabel2.pos".
    /// NOT layout keys (those are handled separately by TryParseDataLabelLayoutKey).
    /// </summary>
    internal static bool TryParseDataLabelDottedKey(string key, out int pointIndex, out string property)
    {
        pointIndex = 0;
        property = "";
        var lower = key.ToLowerInvariant();
        if (!lower.StartsWith("datalabel")) return false;
        var rest = lower["datalabel".Length..];
        var dotIdx = rest.IndexOf('.');
        if (dotIdx <= 0) return false;
        if (!int.TryParse(rest[..dotIdx], out pointIndex) || pointIndex < 1) return false;
        property = rest[(dotIdx + 1)..];
        // Only handle non-layout properties (layout handled by TryParseDataLabelLayoutKey)
        return property is "delete" or "pos" or "position" or "numfmt" or "text";
    }

    internal static void HandleDataLabelDottedProperty(OpenXmlCompositeElement firstSer, int pointIndex, string prop, string value)
    {
        var dLbls = firstSer.GetFirstChild<C.DataLabels>();
        // Auto-create a minimal DataLabels container if not present and we're about to add per-point data.
        if (dLbls == null && (prop == "text" || prop == "delete"))
        {
            dLbls = new C.DataLabels();
            dLbls.AppendChild(new C.ShowLegendKey { Val = false });
            dLbls.AppendChild(new C.ShowValue { Val = true });
            dLbls.AppendChild(new C.ShowCategoryName { Val = false });
            dLbls.AppendChild(new C.ShowSeriesName { Val = false });
            dLbls.AppendChild(new C.ShowPercent { Val = false });
            InsertSeriesChildInOrder(firstSer, dLbls);
        }
        if (dLbls == null) return;

        var ooxmlIdx = (uint)(pointIndex - 1);
        // Coalesce by idx: schema requires at most one <c:dLbl idx="N"> per series.
        // Find-or-create once, then merge subsequent settings into the same element.
        var dLbl = dLbls.Elements<C.DataLabel>()
            .FirstOrDefault(dl => dl.Index?.Val?.Value == ooxmlIdx);
        if (dLbl == null && (prop == "text" || prop == "delete"))
        {
            dLbl = new C.DataLabel();
            dLbl.AppendChild(new C.Index { Val = ooxmlIdx });
            var insertBefore = dLbls.GetFirstChild<C.ShowLegendKey>() as OpenXmlElement
                ?? dLbls.GetFirstChild<C.ShowValue>()
                ?? dLbls.FirstChild;
            if (insertBefore != null) dLbls.InsertBefore(dLbl, insertBefore);
            else dLbls.AppendChild(dLbl);
        }

        switch (prop)
        {
            case "delete":
            {
                if (dLbl == null) return;
                var del = ParseHelpers.IsTruthy(value);
                dLbl.RemoveAllChildren<C.Delete>();
                dLbl.AppendChild(new C.Delete { Val = del });
                // "delete wins" semantics: a deleted label renders nothing, so strip
                // any previously-set visible siblings (tx, numFmt, dLblPos, show*).
                if (del)
                {
                    dLbl.RemoveAllChildren<C.ChartText>();
                    dLbl.RemoveAllChildren<C.NumberingFormat>();
                    dLbl.RemoveAllChildren<C.DataLabelPosition>();
                    dLbl.RemoveAllChildren<C.ShowLegendKey>();
                    dLbl.RemoveAllChildren<C.ShowValue>();
                    dLbl.RemoveAllChildren<C.ShowCategoryName>();
                    dLbl.RemoveAllChildren<C.ShowSeriesName>();
                    dLbl.RemoveAllChildren<C.ShowPercent>();
                    dLbl.RemoveAllChildren<C.ShowBubbleSize>();
                    dLbl.RemoveAllChildren<C.Separator>();
                }
                break;
            }
            case "pos" or "position":
            {
                if (dLbl == null) return;
                // Skip if this dLbl is already marked deleted — delete wins.
                if (dLbl.GetFirstChild<C.Delete>() is { Val.Value: true }) return;
                dLbl.RemoveAllChildren<C.DataLabelPosition>();
                var dlPos = value.ToLowerInvariant() switch
                {
                    "center" or "ctr" => C.DataLabelPositionValues.Center,
                    "insideend" or "inside" => C.DataLabelPositionValues.InsideEnd,
                    "outsideend" or "outside" => C.DataLabelPositionValues.OutsideEnd,
                    "bestfit" or "best" => C.DataLabelPositionValues.BestFit,
                    _ => C.DataLabelPositionValues.OutsideEnd
                };
                dLbl.AppendChild(new C.DataLabelPosition { Val = dlPos });
                break;
            }
            case "numfmt":
            {
                if (dLbl == null) return;
                if (dLbl.GetFirstChild<C.Delete>() is { Val.Value: true }) return;
                dLbl.RemoveAllChildren<C.NumberingFormat>();
                dLbl.AppendChild(new C.NumberingFormat { FormatCode = value, SourceLinked = false });
                break;
            }
            case "text":
            {
                if (dLbl == null) return;
                // Delete wins: if this dLbl is already deleted, ignore a later text= set.
                if (dLbl.GetFirstChild<C.Delete>() is { Val.Value: true }) return;
                dLbl.RemoveAllChildren<C.ChartText>();
                var richText = new C.ChartText();
                var rich = new C.RichText(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.Run(
                            new Drawing.RunProperties { Language = "en-US" },
                            new Drawing.Text(value))));
                richText.AppendChild(rich);
                dLbl.AppendChild(richText);
                // Ensure show flags are present so the custom text renders
                if (dLbl.GetFirstChild<C.ShowValue>() == null)
                    dLbl.AppendChild(new C.ShowValue { Val = true });
                if (dLbl.GetFirstChild<C.ShowCategoryName>() == null)
                    dLbl.AppendChild(new C.ShowCategoryName { Val = false });
                if (dLbl.GetFirstChild<C.ShowSeriesName>() == null)
                    dLbl.AppendChild(new C.ShowSeriesName { Val = false });
                break;
            }
        }

        // Final pass: enforce CT_DLbl schema order. Excel rejects the file silently
        // if children are out of order (Sch_UnexpectedElementContentExpectingComplex).
        // Order: idx, delete, layout, tx, numFmt, spPr, txPr, dLblPos,
        //        showLegendKey, showVal, showCatName, showSerName, showPercent,
        //        showBubbleSize, separator, extLst.
        if (dLbl != null) ReorderDLblChildren(dLbl);
    }

    private static readonly Type[] s_dLblChildOrder =
    {
        typeof(C.Index),
        typeof(C.Delete),
        typeof(C.Layout),
        typeof(C.ChartText),
        typeof(C.NumberingFormat),
        typeof(C.ChartShapeProperties),
        typeof(C.TextProperties),
        typeof(C.DataLabelPosition),
        typeof(C.ShowLegendKey),
        typeof(C.ShowValue),
        typeof(C.ShowCategoryName),
        typeof(C.ShowSeriesName),
        typeof(C.ShowPercent),
        typeof(C.ShowBubbleSize),
        typeof(C.Separator),
        typeof(C.ExtensionList),
    };

    private static void ReorderDLblChildren(C.DataLabel dLbl)
    {
        var kept = new List<OpenXmlElement>();
        foreach (var t in s_dLblChildOrder)
        {
            foreach (var child in dLbl.ChildElements.Where(c => c.GetType() == t).ToList())
            {
                child.Remove();
                kept.Add(child);
            }
        }
        // Re-append in schema order. Any unknown children (shouldn't happen) are dropped.
        foreach (var c in kept) dLbl.AppendChild(c);
    }

    /// <summary>
    /// Parse keys like "legendEntry1.delete".
    /// </summary>
    internal static bool TryParseLegendEntryKey(string key, out int entryIndex)
    {
        entryIndex = 0;
        var lower = key.ToLowerInvariant();
        if (!lower.StartsWith("legendentry")) return false;
        var rest = lower["legendentry".Length..];
        var dotIdx = rest.IndexOf('.');
        if (dotIdx <= 0) return false;
        if (!int.TryParse(rest[..dotIdx], out entryIndex) || entryIndex < 1) return false;
        var prop = rest[(dotIdx + 1)..];
        return prop is "delete" or "hide";
    }

    // ==================== Schema-Order Insertion Helpers ====================

    /// <summary>
    /// Insert a child into a CT_ValAx or CT_CatAx element at the correct schema position.
    /// Schema order (shared prefix): axId, scaling, delete, axPos, majorGridlines, minorGridlines,
    /// title, numFmt, majorTickMark, minorTickMark, tickLblPos, spPr, txPr, crossAx, ...
    /// </summary>
    internal static void InsertAxisChildInOrder(OpenXmlCompositeElement axis, OpenXmlElement child)
    {
        // Elements that come AFTER majorTickMark/minorTickMark/tickLblPos in axis schema
        string[] afterTickElements = ["spPr", "txPr", "crossAx", "crosses", "crossesAt",
            "crossBetween", "auto", "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip",
            "noMultiLvlLbl", "majorUnit", "minorUnit", "dispUnits", "extLst"];

        // For majorTickMark: insert before minorTickMark, tickLblPos, or any afterTickElements
        // For minorTickMark: insert before tickLblPos or any afterTickElements
        // For tickLblPos: insert before spPr, txPr, crossAx, etc.
        string[] insertBeforeNames = child.LocalName switch
        {
            "majorTickMark" => ["minorTickMark", "tickLblPos", ..afterTickElements],
            "minorTickMark" => ["tickLblPos", ..afterTickElements],
            "tickLblPos" => afterTickElements,
            _ => afterTickElements
        };

        foreach (var sibling in axis.ChildElements)
        {
            if (insertBeforeNames.Contains(sibling.LocalName))
            {
                axis.InsertBefore(child, sibling);
                return;
            }
        }
        axis.AppendChild(child);
    }

    /// <summary>
    /// Insert a child into a CT_LineChart at the correct schema position.
    /// Schema: grouping, varyColors, ser+, dLbls, dropLines, hiLowLines, upDownBars, marker, smooth, axId+, extLst
    /// </summary>
    internal static void InsertLineChartChildInOrder(C.LineChart lc, OpenXmlElement child)
    {
        // CT_LineChart schema order: grouping, varyColors, ser*, dLbls?,
        // dropLines?, hiLowLines?, upDownBars?, marker?, smooth?, extLst?, axId+
        string[] insertBeforeNames = child.LocalName switch
        {
            "dropLines" => ["hiLowLines", "upDownBars", "marker", "smooth", "extLst", "axId"],
            "hiLowLines" => ["upDownBars", "marker", "smooth", "extLst", "axId"],
            "upDownBars" => ["marker", "smooth", "extLst", "axId"],
            "marker" => ["smooth", "extLst", "axId"],
            "smooth" => ["extLst", "axId"],
            _ => ["extLst", "axId"]
        };
        foreach (var sibling in lc.ChildElements)
        {
            if (insertBeforeNames.Contains(sibling.LocalName))
            {
                lc.InsertBefore(child, sibling);
                return;
            }
        }
        lc.AppendChild(child);
    }

    /// <summary>
    /// Insert a child into a chart series (CT_*Ser) at the correct schema position.
    /// Common suffix order: ..., dLbls, trendline, errBars, cat/xVal, val/yVal, smooth, extLst
    /// </summary>
    internal static void InsertSeriesChildInOrder(OpenXmlCompositeElement ser, OpenXmlElement child)
    {
        string[] insertBeforeNames = child.LocalName switch
        {
            "dLbls" => ["trendline", "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "bubble3D", "smooth", "extLst"],
            "trendline" => ["errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "bubble3D", "smooth", "extLst"],
            "errBars" => ["cat", "val", "xVal", "yVal", "bubbleSize", "bubble3D", "smooth", "extLst"],
            "smooth" => ["extLst"],
            _ => ["extLst"]
        };

        foreach (var sibling in ser.ChildElements)
        {
            if (insertBeforeNames.Contains(sibling.LocalName))
            {
                ser.InsertBefore(child, sibling);
                return;
            }
        }
        ser.AppendChild(child);
    }

    /// <summary>
    /// Insert effectLst into spPr respecting DrawingML schema: ..., ln, effectLst, effectDag, ...
    /// </summary>
    internal static void InsertEffectListInSpPr(Drawing.ShapeProperties spPr, Drawing.EffectList effectList)
    {
        var ln = spPr.GetFirstChild<Drawing.Outline>();
        if (ln != null) ln.InsertAfterSelf(effectList);
        else spPr.AppendChild(effectList);
    }

    internal static void InsertEffectListInChartSpPr(C.ChartShapeProperties spPr, Drawing.EffectList effectList)
    {
        var ln = spPr.GetFirstChild<Drawing.Outline>();
        if (ln != null) ln.InsertAfterSelf(effectList);
        else spPr.AppendChild(effectList);
    }
}
