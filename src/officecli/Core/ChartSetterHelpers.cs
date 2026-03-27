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
            case "name":
                trendline.RemoveAllChildren<C.TrendlineName>();
                trendline.PrependChild(new C.TrendlineName { Text = value });
                break;
            case "forward":
                trendline.RemoveAllChildren<C.Forward>();
                trendline.AppendChild(new C.Forward { Val = ParseHelpers.SafeParseDouble(value, "trendline.forward") });
                break;
            case "backward":
                trendline.RemoveAllChildren<C.Backward>();
                trendline.AppendChild(new C.Backward { Val = ParseHelpers.SafeParseDouble(value, "trendline.backward") });
                break;
            case "intercept":
                trendline.RemoveAllChildren<C.Intercept>();
                trendline.AppendChild(new C.Intercept { Val = ParseHelpers.SafeParseDouble(value, "trendline.intercept") });
                break;
            case "disprsqr" or "rsquared" or "r2":
                trendline.RemoveAllChildren<C.DisplayRSquaredValue>();
                trendline.AppendChild(new C.DisplayRSquaredValue { Val = ParseHelpers.IsTruthy(value) });
                break;
            case "dispeq" or "equation":
                trendline.RemoveAllChildren<C.DisplayEquation>();
                trendline.AppendChild(new C.DisplayEquation { Val = ParseHelpers.IsTruthy(value) });
                break;
        }
    }

    // ==================== Error Bars Helpers ====================

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
            // Insert before c:dLbls, c:cat, c:val etc.
            var insertBefore = series.GetFirstChild<C.DataLabels>() as OpenXmlElement
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
    /// </summary>
    internal static void HandleSeriesDottedProperty(OpenXmlCompositeElement ser, string prop, string value)
    {
        switch (prop)
        {
            case "smooth":
                ser.RemoveAllChildren<C.Smooth>();
                ser.AppendChild(new C.Smooth { Val = ParseHelpers.IsTruthy(value) });
                break;

            case "trendline":
                ser.RemoveAllChildren<C.Trendline>();
                if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    ser.AppendChild(BuildTrendline(value));
                break;

            case "marker":
                ApplySeriesMarker(ser, value);
                break;

            case "markersize":
            {
                var marker = ser.GetFirstChild<C.Marker>();
                if (marker == null) { marker = new C.Marker(); ser.AppendChild(marker); }
                marker.RemoveAllChildren<C.Size>();
                marker.AppendChild(new C.Size { Val = ParseHelpers.SafeParseByte(value, "series.markerSize") });
                break;
            }

            case "color":
                ApplySeriesColor(ser, value);
                break;

            case "invertifneg" or "invertifnegative":
                ser.RemoveAllChildren<C.InvertIfNegative>();
                ser.AppendChild(new C.InvertIfNegative { Val = ParseHelpers.IsTruthy(value) });
                break;

            case "errbars" or "errorbars":
                ser.RemoveAllChildren<C.ErrorBars>();
                if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    ser.AppendChild(BuildErrorBars(value));
                break;

            case "explosion" or "explode":
                ser.RemoveAllChildren<C.Explosion>();
                if (uint.TryParse(value, out var expVal) && expVal > 0)
                    ser.AppendChild(new C.Explosion { Val = expVal });
                break;

            case "linewidth":
                ApplySeriesLineWidth(ser, (int)(ParseHelpers.SafeParseDouble(value, "series.lineWidth") * 12700));
                break;

            case "linedash" or "dash":
                ApplySeriesLineDash(ser, value);
                break;

            case "shadow":
            {
                var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
                if (spPr == null) { spPr = new C.ChartShapeProperties(); ser.AppendChild(spPr); }
                var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? new Drawing.EffectList();
                if (effectList.Parent == null) spPr.AppendChild(effectList);
                effectList.RemoveAllChildren<Drawing.OuterShadow>();
                if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    effectList.AppendChild(DrawingEffectsHelper.BuildOuterShadow(value, BuildChartColorElement));
                break;
            }

            case "outline":
            {
                var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
                if (spPr == null) { spPr = new C.ChartShapeProperties(); ser.AppendChild(spPr); }
                spPr.RemoveAllChildren<Drawing.Outline>();
                if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    spPr.AppendChild(BuildOutlineElement(value));
                break;
            }

            case "gradient":
                ApplySeriesGradient(ser, value);
                break;

            case "alpha" or "transparency":
            {
                var alphaPercent = ParseHelpers.SafeParseDouble(value, "series.alpha");
                if (prop == "transparency") alphaPercent = 100.0 - alphaPercent;
                ApplySeriesAlpha(ser, (int)(alphaPercent * 1000));
                break;
            }

            default:
                // Trendline sub-properties: series{N}.trendline.name, .forward, .backward, etc.
                if (prop.StartsWith("trendline."))
                {
                    var tl = ser.GetFirstChild<C.Trendline>();
                    if (tl != null)
                        ApplyTrendlineOptions(tl, prop["trendline.".Length..], value);
                    break;
                }
                // Per-point properties: series{N}.point{M}.color, series{N}.point{M}.explosion
                if (prop.StartsWith("point") && TryParsePointKey(prop, out var ptIdx, out var ptProp))
                {
                    switch (ptProp)
                    {
                        case "color":
                            ApplyDataPointColor(ser, ptIdx - 1, value);
                            break;
                        case "explosion" or "explode":
                            ApplyDataPointExplosion(ser, ptIdx - 1,
                                uint.TryParse(value, out var pe) ? pe : 0u);
                            break;
                    }
                }
                break;
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
        // For "text", auto-create a minimal DataLabels container on the series if not present
        if (dLbls == null && prop == "text")
        {
            dLbls = new C.DataLabels();
            dLbls.AppendChild(new C.ShowLegendKey { Val = false });
            dLbls.AppendChild(new C.ShowValue { Val = true });
            dLbls.AppendChild(new C.ShowCategoryName { Val = false });
            dLbls.AppendChild(new C.ShowSeriesName { Val = false });
            dLbls.AppendChild(new C.ShowPercent { Val = false });
            // Insert before AxId or at end of series, per schema order
            var insertBefore = firstSer.GetFirstChild<C.AxisId>() as OpenXmlElement;
            if (insertBefore != null) firstSer.InsertBefore(dLbls, insertBefore);
            else firstSer.AppendChild(dLbls);
        }
        if (dLbls == null) return;

        var ooxmlIdx = (uint)(pointIndex - 1);
        var dLbl = dLbls.Elements<C.DataLabel>()
            .FirstOrDefault(dl => dl.Index?.Val?.Value == ooxmlIdx);

        switch (prop)
        {
            case "delete":
            {
                if (dLbl == null)
                {
                    dLbl = new C.DataLabel();
                    dLbl.Index = new C.Index { Val = ooxmlIdx };
                    dLbl.AppendChild(new C.Delete { Val = ParseHelpers.IsTruthy(value) });
                    var insertBefore = dLbls.GetFirstChild<C.ShowLegendKey>() as OpenXmlElement
                        ?? dLbls.GetFirstChild<C.ShowValue>()
                        ?? dLbls.FirstChild;
                    if (insertBefore != null) dLbls.InsertBefore(dLbl, insertBefore);
                    else dLbls.AppendChild(dLbl);
                }
                else
                {
                    dLbl.RemoveAllChildren<C.Delete>();
                    dLbl.AppendChild(new C.Delete { Val = ParseHelpers.IsTruthy(value) });
                }
                break;
            }
            case "pos" or "position":
            {
                if (dLbl == null) return;
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
                dLbl.RemoveAllChildren<C.NumberingFormat>();
                dLbl.AppendChild(new C.NumberingFormat { FormatCode = value, SourceLinked = false });
                break;
            }
            case "text":
            {
                // Create or find dLbl, then set custom text via c:tx > c:rich
                if (dLbl == null)
                {
                    dLbl = new C.DataLabel();
                    dLbl.Index = new C.Index { Val = ooxmlIdx };
                    var insertBefore = dLbls.GetFirstChild<C.ShowLegendKey>() as OpenXmlElement
                        ?? dLbls.GetFirstChild<C.ShowValue>()
                        ?? dLbls.FirstChild;
                    if (insertBefore != null) dLbls.InsertBefore(dLbl, insertBefore);
                    else dLbls.AppendChild(dLbl);
                }
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
                // Insert tx after idx (schema order: idx, delete, layout, tx, ...)
                var afterIdx = dLbl.GetFirstChild<C.Index>() as OpenXmlElement;
                var afterLayout = dLbl.GetFirstChild<C.Layout>() as OpenXmlElement;
                var insertAfter = afterLayout ?? afterIdx;
                if (insertAfter != null)
                    insertAfter.InsertAfterSelf(richText);
                else
                    dLbl.PrependChild(richText);
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
}
