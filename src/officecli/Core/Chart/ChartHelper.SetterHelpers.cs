// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
        var norm = SchemaKeyNormalizer.Normalize(value);
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
            "linear" => C.TrendlineValues.Linear,
            "exp" or "exponential" => C.TrendlineValues.Exponential,
            "log" or "logarithmic" => C.TrendlineValues.Logarithmic,
            "poly" or "polynomial" => C.TrendlineValues.Polynomial,
            "power" => C.TrendlineValues.Power,
            "movingavg" or "moving" or "movingaverage" => C.TrendlineValues.MovingAverage,
            _ => throw new CliException(
                $"Invalid trendline type '{parts[0]}'. " +
                "Valid: linear, exp, log, poly, power, movingAvg. " +
                "For per-series different trendlines use seriesN.trendline keys, " +
                "not pipe-separated lists.")
                { Code = "invalid_value" }
        };
        trendline.AppendChild(new C.TrendlineType { Val = trendType });

        // Polynomial order or moving average period
        if (parts.Length > 1 && int.TryParse(parts[1], out var order))
        {
            if (trendType == C.TrendlineValues.Polynomial)
                trendline.AppendChild(new C.PolynomialOrder { Val = (byte)Math.Clamp(order, 2, 6) });
            else if (trendType == C.TrendlineValues.MovingAverage)
            {
                // OOXML ST_Skip MinInclusive=2 (c:period inside c:trendline).
                // Pre-fix code silently accepted order=0/1 which Word 422s on.
                if (order < 2)
                    throw new ArgumentException($"movingAvg period must be >= 2 (OOXML ST_Skip MinInclusive=2). Got: {order}.");
                trendline.AppendChild(new C.Period { Val = (uint)order });
            }
            else
            {
                // Treat as forward extrapolation periods
                trendline.AppendChild(new C.Forward { Val = order });
            }
        }
        // OOXML CT_Trendline requires <c:period> when trendlineType=movingAvg;
        // Word rejects the file otherwise. When no explicit period was given,
        // fall back to 2 (Excel's default for "Add Trendline → Moving Average").
        if (trendType == C.TrendlineValues.MovingAverage
            && trendline.GetFirstChild<C.Period>() == null)
        {
            trendline.AppendChild(new C.Period { Val = 2u });
        }
        // Same family for polynomial: <c:order> is required for poly trendlines.
        // Default to degree 2 (Excel's default for "Add Trendline → Polynomial").
        if (trendType == C.TrendlineValues.Polynomial
            && trendline.GetFirstChild<C.PolynomialOrder>() == null)
        {
            trendline.AppendChild(new C.PolynomialOrder { Val = 2 });
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
                // CT_Trendline schema order is:
                //   name → trendlineType → order → period → forward → backward
                //   → intercept → dispRSqr → dispEq → trendlineLbl
                // trendlineLbl is the LAST child. Previous comment had this
                // backwards and the InsertBefore landed tlLbl ahead of
                // trendlineType, which the validator rejected.
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
            {
                // OOXML ST_Skip MinInclusive=2. Pre-fix code clamped < 2 to 2;
                // silently coerce hid invalid input from callers. Throw instead.
                var periodVal = ParseHelpers.SafeParseInt(value, "trendline.period");
                if (periodVal < 2)
                    throw new ArgumentException($"trendline.period must be >= 2 (OOXML ST_Skip MinInclusive=2). Got: {periodVal}.");
                trendline.RemoveAllChildren<C.Period>();
                trendline.AppendChild(new C.Period { Val = (uint)periodVal });
                break;
            }
            case "intercept":
                trendline.RemoveAllChildren<C.Intercept>();
                trendline.AppendChild(new C.Intercept { Val = ParseHelpers.SafeParseDouble(value, "trendline.intercept") });
                break;
            case "disprsqr" or "rsquared" or "r2" or "displayrsquared":
            {
                // CT_Trendline schema order (per ECMA-376 §21.2.2.211):
                //   ... intercept → dispRSqr → dispEq → trendlineLbl → extLst
                // dispRSqr comes BEFORE dispEq. Anchor on the first later-
                // schema sibling so both Set orders produce valid XML.
                trendline.RemoveAllChildren<C.DisplayRSquaredValue>();
                var newRsqr = new C.DisplayRSquaredValue { Val = ParseHelpers.IsTruthy(value) };
                var rsqrAnchor = (OpenXmlElement?)trendline.GetFirstChild<C.DisplayEquation>()
                    ?? trendline.GetFirstChild<C.TrendlineLabel>();
                if (rsqrAnchor != null) trendline.InsertBefore(newRsqr, rsqrAnchor);
                else trendline.AppendChild(newRsqr);
                break;
            }
            case "dispeq" or "equation" or "displayequation":
            {
                // dispEq comes AFTER dispRSqr but BEFORE trendlineLbl per CT_Trendline.
                trendline.RemoveAllChildren<C.DisplayEquation>();
                var newDispEq = new C.DisplayEquation { Val = ParseHelpers.IsTruthy(value) };
                var dispEqAnchor = trendline.GetFirstChild<C.TrendlineLabel>();
                if (dispEqAnchor != null) trendline.InsertBefore(newDispEq, dispEqAnchor);
                else trendline.AppendChild(newDispEq);
                break;
            }
        }
    }

    // ==================== Error Bars Helpers ====================

    /// <summary>
    /// Check if the parent chart type supports errBars on its series (CT_*Ser).
    /// ECMA-376: errBars is a child of CT_LineSer, CT_ScatterSer, CT_BarSer,
    /// CT_AreaSer, CT_BubbleSer (and their 3D variants where applicable).
    /// Not allowed in: pieChart, pie3DChart, doughnutChart, radarChart, stockChart,
    /// surfaceChart, surface3DChart.
    /// </summary>
    internal static bool SeriesSupportsErrorBars(OpenXmlElement ser)
    {
        var parentName = ser.Parent?.LocalName ?? "";
        return parentName is "barChart" or "bar3DChart"
            or "lineChart" or "line3DChart"
            or "scatterChart"
            or "areaChart" or "area3DChart"
            or "bubbleChart";
    }

    // Single source of truth for labelPos alias parsing. Accepts every
    // friendly alias plus the raw schema tokens the Reader emits verbatim
    // (ctr, t, b, l, r, outEnd, inEnd, inBase, bestFit) so dump→batch replay
    // always parses. Unknown tokens throw instead of silently coercing to
    // OutsideEnd (silent-accept enum-miss family); the three former inline
    // switches each covered a different subset, so a token accepted on one
    // path could throw or coerce on another.
    internal static C.DataLabelPositionValues ParseDataLabelPosition(string value) =>
        value.ToLowerInvariant() switch
        {
            "center" or "ctr" => C.DataLabelPositionValues.Center,
            "insideend" or "inside" or "inend" => C.DataLabelPositionValues.InsideEnd,
            "outsideend" or "outside" or "outend" => C.DataLabelPositionValues.OutsideEnd,
            "insidebase" or "inbase" or "base" => C.DataLabelPositionValues.InsideBase,
            "top" or "t" => C.DataLabelPositionValues.Top,
            "bottom" or "b" => C.DataLabelPositionValues.Bottom,
            "left" or "l" => C.DataLabelPositionValues.Left,
            "right" or "r" => C.DataLabelPositionValues.Right,
            "bestfit" or "best" => C.DataLabelPositionValues.BestFit,
            _ => throw new ArgumentException(
                $"Unknown label position '{value}'. Valid: center, insideEnd, outsideEnd, insideBase, top, bottom, left, right, bestFit.")
        };

    internal static C.ErrorBars BuildErrorBars(string spec)
    {
        // Format: "type" or "type:value" e.g. "fixed:5", "percent:10", "stddev", "stderr"
        // R55 bt-6: cust spec is "cust:<direction>:<plusCSV>:<minusCSV>" — emit
        // per-direction <c:plus>/<c:minus> NumberLiteral arrays. The Reader
        // pairs with this form (ReadErrorBarSideCsv) so dump-replay of cust
        // error bars round-trips through the inline numLit cache, not numRef
        // (which would need a live cell reference on the embedded workbook).
        // CONSISTENCY(errorbars-bare-number): bare number (e.g. "5") is taken as
        // fixed:<N>, mirroring how other chart numeric specs accept a value-only
        // shorthand. Without this, "5" matched no type-name arm and fell through
        // to FixedValue with no magnitude — producing zero-height error bars.
        var parts = spec.Split(':');
        var typeStr = parts[0].Trim().ToLowerInvariant();
        string? bareValue = null;
        if (parts.Length == 1 && double.TryParse(typeStr,
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out _))
        {
            bareValue = typeStr;
            typeStr = "fixed";
        }

        // Direction-keyword leading form: errBars=both | plus | minus | both:fixed:5
        // The first slot is the ErrorBarType direction, not the value-type. Shift
        // remaining slots so the existing typeStr / value parsing picks them up.
        // Bare direction (no second slot) defaults to type=stdErr — Excel's UI
        // default for "Error Bars > Standard Error" with direction=Both.
        var explicitDirection = (C.ErrorBarValues?)null;
        if (typeStr is "both" or "plus" or "minus")
        {
            explicitDirection = typeStr switch
            {
                "plus" => C.ErrorBarValues.Plus,
                "minus" => C.ErrorBarValues.Minus,
                _ => C.ErrorBarValues.Both,
            };
            // Shift: parts[1] becomes the type, parts[2..] becomes value(s).
            // If no second slot, fall back to stdErr (no magnitude needed).
            typeStr = parts.Length > 1 ? parts[1].Trim().ToLowerInvariant() : "stderr";
            parts = parts.Length > 1
                ? new[] { typeStr }.Concat(parts.Skip(2)).ToArray()
                : new[] { typeStr };
        }

        var errBars = new C.ErrorBars();
        errBars.AppendChild(new C.ErrorDirection { Val = C.ErrorBarDirectionValues.Y });

        // R55 bt-6: cust path. parts = ["cust", direction, plusCSV, minusCSV].
        if (typeStr == "cust" && parts.Length >= 4)
        {
            var direction = (parts[1].Trim().ToLowerInvariant()) switch
            {
                "plus" => C.ErrorBarValues.Plus,
                "minus" => C.ErrorBarValues.Minus,
                _ => C.ErrorBarValues.Both
            };
            errBars.AppendChild(new C.ErrorBarType { Val = direction });
            errBars.AppendChild(new C.ErrorBarValueType { Val = C.ErrorValues.Custom });

            static C.NumberLiteral BuildLit(string csv)
            {
                var lit = new C.NumberLiteral(new C.FormatCode("General"));
                var values = csv.Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(v => v.Trim()).ToList();
                lit.AppendChild(new C.PointCount { Val = (uint)values.Count });
                uint idx = 0;
                foreach (var v in values)
                {
                    lit.AppendChild(new C.NumericPoint(new C.NumericValue(v)) { Index = idx++ });
                }
                return lit;
            }

            // Schema order: <c:plus> before <c:minus>.
            if (!string.IsNullOrEmpty(parts[2]))
                errBars.AppendChild(new C.Plus(BuildLit(parts[2])));
            if (!string.IsNullOrEmpty(parts[3]))
                errBars.AppendChild(new C.Minus(BuildLit(parts[3])));
            return errBars;
        }

        errBars.AppendChild(new C.ErrorBarType { Val = explicitDirection ?? C.ErrorBarValues.Both });

        var errValType = typeStr switch
        {
            "fixed" or "fixedvalue" => C.ErrorValues.FixedValue,
            "percent" or "pct" or "percentage" => C.ErrorValues.Percentage,
            "stddev" or "standarddeviation" => C.ErrorValues.StandardDeviation,
            "stderr" or "standarderror" => C.ErrorValues.StandardError,
            // Unknown token must fail loudly: "std" silently coerced to
            // FixedValue produced zero/wrong error bars with no warning
            // (silent-accept enum-miss family).
            _ => throw new ArgumentException(
                $"Unknown error-bar type '{typeStr}'. Valid: fixed[:N], percent[:N], stddev[:N], stderr, " +
                $"cust:<direction>:<plusCSV>:<minusCSV>, optionally prefixed with both:/plus:/minus:.")
        };
        errBars.AppendChild(new C.ErrorBarValueType { Val = errValType });

        var magnitudeStr = bareValue ?? (parts.Length > 1 ? parts[1] : null);
        if (magnitudeStr != null && double.TryParse(magnitudeStr,
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

    /// <summary>
    /// a:ln/@w schema ceiling (ST_LineWidth MaxInclusive): 20116800 EMU = 1584pt.
    /// </summary>
    internal const int MaxLineWidthEmu = 20116800;

    /// <summary>
    /// Parse one line-width token (the width slot of "color:width[:dash]" specs
    /// and the dotted *.width keys) into EMU for a:ln/@w.
    /// - Bare numbers are POINTS — the documented colon-spec unit
    ///   (schemas/help: "color:width", examples 0.5 / 1 / 1.5).
    /// - Unit-qualified values ("1pt", "0.5mm", "0.02in", "12700emu") go through
    ///   EmuConverter.ParseEmu.
    /// - Bare integers too large to be a legal point width (&gt; 1584pt) are RAW
    ///   EMU — the ParseEmu raw-integer convention. Width values copied out of
    ///   real OOXML (e.g. "…:12700" = 1pt) must round-trip as-is instead of
    ///   being re-multiplied by 12700 into schema-invalid XML.
    /// The result is clamped to [0, MaxLineWidthEmu] so Set never emits an
    /// a:ln/@w that fails OOXML validation.
    /// </summary>
    internal static bool TryParseLineWidthEmu(string? token, out int emu)
    {
        emu = 0;
        token = token?.Trim();
        if (string.IsNullOrEmpty(token)) return false;
        if (char.IsLetter(token[^1]))
        {
            try { emu = (int)Math.Clamp(EmuConverter.ParseEmu(token), 0, MaxLineWidthEmu); }
            catch (ArgumentException) { return false; }
            return true;
        }
        if (!double.TryParse(token, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var num)
            || double.IsNaN(num) || double.IsInfinity(num) || num < 0)
            return false;
        var asPointsEmu = num * EmuConverter.EmuPerPointF;
        emu = asPointsEmu > MaxLineWidthEmu && num == Math.Floor(num)
            ? (int)Math.Min(num, MaxLineWidthEmu)   // raw EMU integer
            : (int)Math.Min(asPointsEmu, MaxLineWidthEmu);
        return true;
    }

    internal static Drawing.Outline BuildOutlineElement(string spec)
    {
        // Format: "color" or "color:width" or "color:width:dash"
        // e.g. "000000", "333333:1.5", "666666:1:dash"
        var parts = spec.Split(':');
        var color = parts[0].Trim();
        var widthEmu = parts.Length > 1 && TryParseLineWidthEmu(parts[1], out var w)
            ? w : (int)(0.75 * EmuConverter.EmuPerPoint);
        var dash = parts.Length > 2 ? parts[2].Trim() : null;

        var outline = new Drawing.Outline { Width = widthEmu };
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
            // Route through the shared anchor helper so dPt lands after
            // marker and before dLbls/trendline/errBars/cat/val/xVal/yVal/
            // bubbleSize/smooth. Anchoring only on Values/CategoryAxisData
            // misses scatter/bubble series (data is in xVal/yVal/bubbleSize),
            // so dPt was appended after the data tail — the validator then
            // reports "unexpected child element 'c:dPt'".
            InsertSeriesChildInOrder(series, dPt);
        }

        var spPr = dPt.GetFirstChild<C.ChartShapeProperties>();
        if (spPr == null) { spPr = new C.ChartShapeProperties(); dPt.AppendChild(spPr); }
        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.NoFill>();
        if (color.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            spPr.PrependChild(new Drawing.NoFill());
            return;
        }
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
            InsertSeriesChildInOrder(series, dPt);
        }
        dPt.RemoveAllChildren<C.Explosion>();
        if (explosion > 0)
            dPt.AppendChild(new C.Explosion { Val = explosion });
    }

    /// <summary>
    /// Get-or-create the <c:dPt> for the given 0-based point index and
    /// position it within the schema-correct slot of the series. Returns
    /// the existing element when present.
    /// </summary>
    private static C.DataPoint EnsureDataPoint(OpenXmlCompositeElement series, int pointIndex)
    {
        var dPts = series.Elements<C.DataPoint>().ToList();
        var dPt = dPts.FirstOrDefault(dp => dp.Index?.Val?.Value == (uint)pointIndex);
        if (dPt != null) return dPt;
        dPt = new C.DataPoint();
        dPt.AppendChild(new C.Index { Val = (uint)pointIndex });
        InsertSeriesChildInOrder(series, dPt);
        return dPt;
    }

    /// <summary>
    /// Insert <paramref name="child"/> into the dPt in CT_DPt schema order:
    /// idx, invertIfNegative, marker, bubble3D, explosion, dLbl, spPr, extLst.
    /// </summary>
    private static void InsertDataPointChildInOrder(C.DataPoint dPt, OpenXmlElement child)
    {
        // Find first existing child whose schema rank is strictly greater
        // than the new child's rank; insert before it.
        int Rank(OpenXmlElement el) => el switch
        {
            C.Index => 0,
            C.InvertIfNegative => 1,
            C.Marker => 2,
            C.Bubble3D => 3,
            C.Explosion => 4,
            C.DataLabel => 5,
            C.ChartShapeProperties => 6,
            C.ExtensionList => 7,
            _ => 99,
        };
        var newRank = Rank(child);
        OpenXmlElement? anchor = null;
        foreach (var existing in dPt.ChildElements)
        {
            if (Rank(existing) > newRank) { anchor = existing; break; }
        }
        if (anchor != null) dPt.InsertBefore(child, anchor);
        else dPt.AppendChild(child);
    }

    internal static bool ApplyDataPointMarker(OpenXmlCompositeElement series, int pointIndex, string markerSpec)
    {
        // Mirror ApplySeriesMarker (style[:size[:color]]) but scope to the
        // matching <c:dPt>. Reject unknown style tokens up the call chain so
        // callers surface UNSUPPORTED instead of silent corruption.
        var parts = markerSpec.Split(':');
        var styleToken = parts[0].Trim().ToLowerInvariant();
        C.MarkerStyleValues style;
        switch (styleToken)
        {
            case "circle":   style = C.MarkerStyleValues.Circle; break;
            case "diamond":  style = C.MarkerStyleValues.Diamond; break;
            case "square":   style = C.MarkerStyleValues.Square; break;
            case "triangle": style = C.MarkerStyleValues.Triangle; break;
            case "star":     style = C.MarkerStyleValues.Star; break;
            case "x":        style = C.MarkerStyleValues.X; break;
            case "plus":     style = C.MarkerStyleValues.Plus; break;
            case "dash":     style = C.MarkerStyleValues.Dash; break;
            case "dot":      style = C.MarkerStyleValues.Dot; break;
            case "none":     style = C.MarkerStyleValues.None; break;
            case "auto":     style = C.MarkerStyleValues.Auto; break;
            default:         return false;
        }
        var dPt = EnsureDataPoint(series, pointIndex);
        var existing = dPt.GetFirstChild<C.Marker>();
        var existingSize = existing?.GetFirstChild<C.Size>()?.CloneNode(true) as C.Size;
        var existingSpPr = existing?.GetFirstChild<C.ChartShapeProperties>()?.CloneNode(true) as C.ChartShapeProperties;
        dPt.RemoveAllChildren<C.Marker>();
        var marker = new C.Marker();
        marker.AppendChild(new C.Symbol { Val = style });
        if (parts.Length > 1 && byte.TryParse(parts[1], out var size))
            marker.AppendChild(new C.Size { Val = size });
        else if (existingSize != null)
            marker.AppendChild(existingSize);
        if (parts.Length > 2)
        {
            var mSpPr = new C.ChartShapeProperties();
            var fill = new Drawing.SolidFill();
            fill.AppendChild(BuildChartColorElement(parts[2]));
            mSpPr.AppendChild(fill);
            marker.AppendChild(mSpPr);
        }
        else if (existingSpPr != null)
            marker.AppendChild(existingSpPr);
        InsertDataPointChildInOrder(dPt, marker);
        return true;
    }

    internal static bool ApplyDataPointMarkerSize(OpenXmlCompositeElement series, int pointIndex, string value)
    {
        if (!byte.TryParse(value, out var size)) return false;
        var dPt = EnsureDataPoint(series, pointIndex);
        var marker = dPt.GetFirstChild<C.Marker>();
        if (marker == null)
        {
            marker = new C.Marker();
            InsertDataPointChildInOrder(dPt, marker);
        }
        marker.RemoveAllChildren<C.Size>();
        // CT_Marker order: symbol, size, spPr, extLst — Size must follow Symbol.
        var symbol = marker.GetFirstChild<C.Symbol>();
        var sizeEl = new C.Size { Val = size };
        if (symbol != null) symbol.InsertAfterSelf(sizeEl);
        else marker.PrependChild(sizeEl);
        return true;
    }

    internal static bool ApplyDataPointMarkerColor(OpenXmlCompositeElement series, int pointIndex, string color)
    {
        var dPt = EnsureDataPoint(series, pointIndex);
        var marker = dPt.GetFirstChild<C.Marker>();
        if (marker == null)
        {
            marker = new C.Marker();
            InsertDataPointChildInOrder(dPt, marker);
        }
        var mSpPr = marker.GetFirstChild<C.ChartShapeProperties>();
        if (mSpPr == null)
        {
            mSpPr = new C.ChartShapeProperties();
            // CT_Marker schema order: symbol, size, spPr, extLst.
            var anchor = marker.GetFirstChild<C.ExtensionList>() as OpenXmlElement;
            if (anchor != null) marker.InsertBefore(mSpPr, anchor);
            else marker.AppendChild(mSpPr);
        }
        mSpPr.RemoveAllChildren<Drawing.SolidFill>();
        var fill = new Drawing.SolidFill();
        fill.AppendChild(BuildChartColorElement(color));
        mSpPr.PrependChild(fill);
        return true;
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
                //
                // R28-B2: Reader emits semicolon-joined spec list when a
                // series carries multiple trendlines (e.g. "linear;poly:3").
                // Split here so dump→replay re-applies each; single-spec
                // input (no ';') still hits the legacy append-or-replace
                // path unchanged.
                if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    ser.RemoveAllChildren<C.Trendline>();
                }
                else
                {
                    var specs = value.Contains(';')
                        ? value.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        : new[] { value };
                    foreach (var spec in specs)
                    {
                        var newTl = BuildTrendline(spec);
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
                }
                return true;

            case "marker":
                // Return ApplySeriesMarker's result directly — false propagates
                // unsupported up so callers (HandleSeriesDottedProperty contract)
                // surface marker= with an invalid token as UNSUPPORTED.
                return ApplySeriesMarker(ser, value);

            case "markersize":
            case "marker.size":
            {
                var marker = ser.GetFirstChild<C.Marker>();
                if (marker == null)
                {
                    marker = new C.Marker();
                    // CONSISTENCY(insert-series-child): route via the shared
                    // helper so a markerSize set after point.color (dPt) or
                    // dLbls still lands at the schema-correct position.
                    InsertSeriesChildInOrder(ser, marker);
                }
                marker.RemoveAllChildren<C.Size>();
                // CT_Marker order: symbol, size, spPr, extLst — Size must follow
                // Symbol and precede spPr. AppendChild landed it after spPr when
                // markerColor ran first (the dump emits marker→markerColor→
                // markerSize), producing 'unexpected child element size'. Mirrors
                // ApplyDataPointMarkerSize.
                var szEl = new C.Size { Val = ParseHelpers.SafeParseByte(value, "series.markerSize") };
                var sym = marker.GetFirstChild<C.Symbol>();
                if (sym != null) sym.InsertAfterSelf(szEl);
                else marker.PrependChild(szEl);
                return true;
            }

            case "markercolor":
            case "marker.color":
            {
                // CONSISTENCY(marker-dotted): mirror markersize — write fill on
                // the existing marker's spPr, preserving symbol/size so the
                // dumped marker= / markerSize= / markerColor= triplet round-trips
                // without one key clobbering another.
                var existing = ser.GetFirstChild<C.Marker>();
                var existingSym = existing?.GetFirstChild<C.Symbol>()?.CloneNode(true) as C.Symbol;
                var existingSize = existing?.GetFirstChild<C.Size>()?.CloneNode(true) as C.Size;
                if (existing != null) existing.Remove();
                var marker = new C.Marker();
                if (existingSym != null) marker.AppendChild(existingSym);
                else marker.AppendChild(new C.Symbol { Val = C.MarkerStyleValues.Circle });
                if (existingSize != null) marker.AppendChild(existingSize);
                var mSpPr = new C.ChartShapeProperties();
                var fill = new Drawing.SolidFill();
                fill.AppendChild(BuildChartColorElement(value));
                mSpPr.AppendChild(fill);
                marker.AppendChild(mSpPr);
                // CONSISTENCY(insert-series-child): see ApplySeriesMarker —
                // route through the shared helper to pick up the full marker
                // anchor list (including dPt/dLbls).
                InsertSeriesChildInOrder(ser, marker);
                return true;
            }

            case "marker.style":
            {
                // CONSISTENCY(marker-dotted): mirror "marker=circle" but accept the
                // dotted alternative seriesN.marker.style=circle. Preserve any
                // existing c:size so users can set style and size independently.
                // Returns false (unsupported) for invalid tokens like `picture`.
                var existing = ser.GetFirstChild<C.Marker>();
                var existingSize = existing?.GetFirstChild<C.Size>()?.Val?.Value;
                if (!ApplySeriesMarker(ser, value)) return false;
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
                    // ATOMICITY: build (and fully validate) the replacement
                    // content BEFORE touching the existing <c:val>. The old code
                    // called RemoveAllChildren() first, so a rejected token
                    // (e.g. a named range like "SalesRange", which is neither a
                    // Sheet!ref nor a number list) threw AFTER emptying <c:val>,
                    // leaving a schema-invalid empty element that real Excel
                    // refuses to open (0x800A03EC) — and it persisted on save
                    // even though the CLI reported an error.
                    C.Values builtVals;
                    if (value.Contains('!'))
                    {
                        // Cell reference: e.g. Sheet1!B2:B4 — normalize so a
                        // sheet name that needs quoting (spaces, hyphens, leading
                        // digit) is quoted, matching the Add path.
                        builtVals = BuildValuesRef(NormalizeRangeReference(value));
                    }
                    else
                    {
                        // Mirror the Add path's ParseSeriesValues guard. The
                        // old TryParse ? d : 0.0 fallback silently coerced
                        // NaN / Infinity / unparsable tokens into 0.0, and
                        // worse, NaN / Infinity that *did* parse went straight
                        // into <c:v> as text — OOXML <c:v> requires a finite
                        // double. Reject non-finite values and empty tokens
                        // with invalid_value so set chart values=NaN behaves
                        // the same as add chart values=NaN.
                        var tokens = value.Split(',');
                        var nums = new double[tokens.Length];
                        for (int ti = 0; ti < tokens.Length; ti++)
                        {
                            var trimmed = tokens[ti].Trim();
                            if (string.IsNullOrEmpty(trimmed))
                                throw new CliException(
                                    $"values: empty token at position {ti + 1}. Expected comma-separated finite numbers (e.g. '1,2,3').")
                                    { Code = "invalid_value" };
                            if (!double.TryParse(trimmed, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var d)
                                || double.IsNaN(d) || double.IsInfinity(d))
                                throw new CliException(
                                    $"values: invalid number '{trimmed}'. Expected comma-separated finite numbers (e.g. '1,2,3').")
                                    { Code = "invalid_value" };
                            nums[ti] = d;
                        }
                        builtVals = BuildValues(nums);
                    }
                    // Validation passed — now safe to swap the content in.
                    valEl.RemoveAllChildren();
                    foreach (var child in builtVals.ChildElements.ToList())
                        valEl.AppendChild(child.CloneNode(true));
                }
                return true;
            }

            case "invertifneg" or "invertifnegative":
                ser.RemoveAllChildren<C.InvertIfNegative>();
                // CT_BarSer/CT_AreaSer/CT_PieSer all require invertIfNegative
                // immediately after spPr — append would drop it past dPt/dLbls
                // and render as no-op (see InsertSeriesChildInOrder doc).
                InsertSeriesChildInOrder(ser, new C.InvertIfNegative { Val = ParseHelpers.IsTruthy(value) });
                return true;

            // BUG-DUMP-R33-1: inject the Reader's verbatim styling fragments.
            // value is the captured OuterXml; parse it back into the typed SDK
            // element and splice it in at the schema-correct position. The
            // fragments carry their own a:/c: namespace declarations (the
            // Reader takes OuterXml from live elements), so the SDK
            // string-constructor round-trips them losslessly. Existing
            // same-name children are removed first so a Set-after-Set (or the
            // ApplySeriesColor that the column/bar Builder runs before deferred
            // props apply) doesn't leave a duplicate.
            case "sppr":
            {
                if (string.IsNullOrWhiteSpace(value)) return true;
                ser.RemoveAllChildren<C.ChartShapeProperties>();
                InsertSeriesChildInOrder(ser, new C.ChartShapeProperties(value));
                return true;
            }

            case "dpt":
            {
                if (string.IsNullOrWhiteSpace(value)) return true;
                ser.RemoveAllChildren<C.DataPoint>();
                // \x1e (record separator) joins the per-point fragments on the
                // Reader side; never appears inside XML, so the split is safe.
                foreach (var frag in value.Split('\x1e', StringSplitOptions.RemoveEmptyEntries))
                    InsertSeriesChildInOrder(ser, new C.DataPoint(frag));
                return true;
            }

            case "dlbls":
            {
                if (string.IsNullOrWhiteSpace(value)) return true;
                ser.RemoveAllChildren<C.DataLabels>();
                InsertSeriesChildInOrder(ser, new C.DataLabels(value));
                return true;
            }

            case "errbars" or "errorbars":
                ser.RemoveAllChildren<C.ErrorBars>();
                if (!value.Equals("none", StringComparison.OrdinalIgnoreCase)
                    && SeriesSupportsErrorBars(ser))
                    InsertSeriesChildInOrder(ser, BuildErrorBars(value));
                return true;

            case "explosion" or "explode":
                ser.RemoveAllChildren<C.Explosion>();
                if (uint.TryParse(value, out var expVal) && expVal > 0)
                    InsertSeriesChildInOrder(ser, new C.Explosion { Val = expVal });
                return true;

            case "linewidth":
            case "outlinewidth":
                if (TryParseLineWidthEmu(value, out var lnWidthEmu))
                    ApplySeriesLineWidth(ser, lnWidthEmu);
                else
                    // Preserve the structured invalid_value error for garbage input.
                    ParseHelpers.SafeParseDouble(value, "series.lineWidth");
                return true;

            case "linedash" or "dash":
            case "outlinedash":
                ApplySeriesLineDash(ser, value);
                return true;

            case "outlinecolor":
            case "linecolor":
            {
                // Reader emits per-series outline as separate keys
                // (outlineColor, lineWidth, lineDash). The existing `outline`
                // setter takes a compound `color:width:dash` spec and would
                // require callers to round-trip via the compound form. Accept
                // the read-side names directly so dump→batch replays one prop
                // per emit. Update only the SolidFill child; preserve any
                // existing width / dash on the outline element.
                // Route through the schema-aware helper — appending spPr at
                // the end of CT_ScatterSer / CT_LineSer breaks the required
                // child-element order (idx, order, tx, spPr, marker, …) and
                // PowerPoint rejects the file (Error 0x80070570 / "needs
                // repair"). CONSISTENCY(chart-schema-order).
                var spPr = GetOrCreateSeriesShapeProperties(ser);
                var ln = spPr.GetFirstChild<Drawing.Outline>();
                if (ln == null)
                {
                    ln = new Drawing.Outline();
                    var effLst = spPr.GetFirstChild<Drawing.EffectList>();
                    if (effLst != null) spPr.InsertBefore(ln, effLst);
                    else spPr.AppendChild(ln);
                }
                // BuildScatterChart seeds marker-only series with <a:ln><a:noFill/></a:ln>
                // to suppress connecting lines. A subsequent outlineColor / lineWidth
                // / lineDash write must drop NoFill — otherwise <a:ln> ends up with
                // both NoFill AND SolidFill, which is schema-invalid and trips
                // PowerPoint "repair" (Error 422 on open).
                ln.RemoveAllChildren<Drawing.NoFill>();
                ln.RemoveAllChildren<Drawing.SolidFill>();
                var newFill = new Drawing.SolidFill();
                newFill.AppendChild(BuildChartColorElement(value));
                // SolidFill must precede PrstDash inside a:ln per schema.
                var prstDashEl = ln.GetFirstChild<Drawing.PresetDash>();
                if (prstDashEl != null) ln.InsertBefore(newFill, prstDashEl);
                else ln.PrependChild(newFill);
                return true;
            }

            case "shadow":
            {
                var spPr = GetOrCreateSeriesShapeProperties(ser);
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
                var spPr = GetOrCreateSeriesShapeProperties(ser);
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
                // Per-series labelFont (compound + dotted) — mirrors the chart-level
                // labelFont fan-out (Setter.cs cases `labelfont` / `labelfont.*`)
                // but scoped to this series' own <c:dLbls>. Without this dispatch,
                // `series{N}.labelFont*=` fell through to the catch-all `return
                // false;` and surfaced as UNSUPPORTED even though the chart-level
                // form worked. dLbls is created on the series if absent so the
                // first labelFont set is non-destructive.
                if (prop.Equals("labelfont", StringComparison.OrdinalIgnoreCase))
                {
                    var dl = EnsureSeriesDataLabels(ser);
                    dl.RemoveAllChildren<C.TextProperties>();
                    dl.PrependChild(BuildLabelTextProperties(value));
                    return true;
                }
                if (prop.StartsWith("labelfont.", StringComparison.OrdinalIgnoreCase))
                {
                    var subkey = prop.Substring("labelfont.".Length).ToLowerInvariant();
                    if (subkey is "color" or "size" or "bold" or "name" or "font")
                    {
                        var dl = EnsureSeriesDataLabels(ser);
                        ApplyLabelFontSubkey(dl, subkey, value);
                        return true;
                    }
                    return false;
                }
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
                        // R38: per-point marker / markerSize / markerColor —
                        // ApplySeriesMarker writes a <c:marker> on the <c:ser>
                        // element; the data-point equivalent writes the same
                        // <c:marker> child under the matching <c:dPt>. Reuse
                        // the spec parser by routing through a per-point
                        // applier that mirrors ApplySeriesMarker but scopes
                        // to the dPt.
                        case "marker":
                            return ApplyDataPointMarker(ser, ptIdx - 1, value);
                        case "markersize":
                            return ApplyDataPointMarkerSize(ser, ptIdx - 1, value);
                        case "markercolor":
                            return ApplyDataPointMarkerColor(ser, ptIdx - 1, value);
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

    /// <summary>
    /// Get-or-create a series-scoped &lt;c:dLbls&gt; container with the
    /// minimal show* skeleton (mirrors EnsureDataLabelsOnAllChartGroups but
    /// per-series). Used by per-series labelFont fan-out — the chart-level
    /// EnsureDataLabelsOnAllChartGroups attaches dLbls under the chart-group,
    /// not under each &lt;c:ser&gt;.
    /// </summary>
    private static C.DataLabels EnsureSeriesDataLabels(OpenXmlCompositeElement ser)
    {
        var existing = ser.GetFirstChild<C.DataLabels>();
        if (existing != null) return existing;
        var dLbls = new C.DataLabels();
        dLbls.AppendChild(new C.ShowLegendKey { Val = false });
        dLbls.AppendChild(new C.ShowValue { Val = false });
        dLbls.AppendChild(new C.ShowCategoryName { Val = false });
        dLbls.AppendChild(new C.ShowSeriesName { Val = false });
        dLbls.AppendChild(new C.ShowPercent { Val = false });
        dLbls.AppendChild(new C.ShowBubbleSize { Val = false });
        InsertSeriesChildInOrder(ser, dLbls);
        return dLbls;
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
                dLbl.AppendChild(new C.DataLabelPosition { Val = ParseDataLabelPosition(value) });
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
        // Elements that come AFTER axPos in the shared axis prefix
        // (axId, scaling, delete, axPos, majorGridlines, minorGridlines, title,
        // numFmt, majorTickMark, minorTickMark, tickLblPos, ...afterTickElements).
        string[] afterAxPos = ["majorGridlines", "minorGridlines", "title", "numFmt",
            "majorTickMark", "minorTickMark", "tickLblPos", ..afterTickElements];

        // For axPos: insert before majorGridlines and everything after.
        // For majorTickMark: insert before minorTickMark, tickLblPos, or any afterTickElements
        // For minorTickMark: insert before tickLblPos or any afterTickElements
        // For tickLblPos: insert before spPr, txPr, crossAx, etc.
        // CONSISTENCY(catax-tail-order): CT_CatAx-only tail elements (auto,
        // lblAlgn, lblOffset, tickLblSkip, tickMarkSkip, noMultiLvlLbl) come
        // AFTER crossAx/crosses/crossesAt/crossBetween. The generic `_` fallback
        // would otherwise anchor them before crossAx and produce an invalid file
        // ("unexpected lblOffset, expected crossAx").
        string[] insertBeforeNames = child.LocalName switch
        {
            "axPos" => afterAxPos,
            "majorTickMark" => ["minorTickMark", "tickLblPos", ..afterTickElements],
            "minorTickMark" => ["tickLblPos", ..afterTickElements],
            "tickLblPos" => afterTickElements,
            "auto" => ["lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "extLst"],
            "lblAlgn" => ["lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "extLst"],
            "lblOffset" => ["tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "extLst"],
            "tickLblSkip" => ["tickMarkSkip", "noMultiLvlLbl", "extLst"],
            "tickMarkSkip" => ["noMultiLvlLbl", "extLst"],
            "noMultiLvlLbl" => ["extLst"],
            "majorUnit" => ["minorUnit", "dispUnits", "extLst"],
            "minorUnit" => ["dispUnits", "extLst"],
            "dispUnits" => ["extLst"],
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
    /// Insert a <c>&lt;c:dLbls&gt;</c> element into a chart-group element
    /// (CT_BarChart / CT_LineChart / CT_PieChart / CT_ScatterChart / etc.) at
    /// the correct schema position. dLbls schema-orders BEFORE the optional
    /// per-group tail (dropLines, hiLowLines, upDownBars, gapWidth, overlap,
    /// showMarker, holeSize, firstSliceAngle) and the mandatory axId(+).
    ///
    /// CONSISTENCY(insert-chart-group-dlbls): callers used to hand-roll the
    /// same anchor chain three times (datalabels= bootstrap, labelPos=
    /// bootstrap, datalabels.show* bootstrap) and one of them used
    /// PrependChild which lands dLbls before barDir/ser — schema-invalid.
    /// Centralized here so future chart-group dLbls insertions get the right
    /// position without re-deriving the chain.
    /// </summary>
    internal static void InsertChartGroupDLbls(OpenXmlElement chartGroup, C.DataLabels dLbls)
    {
        var anchor = chartGroup.GetFirstChild<C.DropLines>() as OpenXmlElement
            ?? chartGroup.GetFirstChild<C.HighLowLines>() as OpenXmlElement
            ?? chartGroup.GetFirstChild<C.UpDownBars>() as OpenXmlElement
            ?? chartGroup.GetFirstChild<C.GapWidth>() as OpenXmlElement
            ?? chartGroup.GetFirstChild<C.Overlap>() as OpenXmlElement
            ?? chartGroup.GetFirstChild<C.ShowMarker>() as OpenXmlElement
            ?? chartGroup.GetFirstChild<C.HoleSize>() as OpenXmlElement
            ?? chartGroup.GetFirstChild<C.FirstSliceAngle>() as OpenXmlElement
            ?? (OpenXmlElement?)chartGroup.GetFirstChild<C.AxisId>();
        if (anchor != null) chartGroup.InsertBefore(dLbls, anchor);
        else chartGroup.AppendChild(dLbls);
    }

    /// <summary>
    /// Insert a child into a CT_LineChart at the correct schema position.
    /// Schema: grouping, varyColors, ser+, dLbls, dropLines, hiLowLines, upDownBars, marker, smooth, axId+, extLst
    /// </summary>
    internal static void InsertLineChartChildInOrder(OpenXmlCompositeElement lc, OpenXmlElement child)
    {
        // CT_LineChart schema order: grouping, varyColors, ser*, dLbls?,
        // dropLines?, hiLowLines?, upDownBars?, marker?, smooth?, extLst?, axId+
        // CT_StockChart (ser+, dLbls?, dropLines?, hiLowLines?, upDownBars?,
        // axId+) is a strict subsequence, so the same anchor chain serves both.
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
    /// Schema order (CT_BarSer is the strictest; other Ser types are subsequences):
    ///   idx, order, tx, spPr, invertIfNegative, pictureOptions, dPt, dLbls,
    ///   trendline, errBars, cat, val, xVal, yVal, bubbleSize, bubble3D, shape,
    ///   smooth, extLst
    /// PowerPoint silently drops out-of-order children (e.g. invertIfNegative
    /// appended after dLbls renders as if absent — negative bars stay un-inverted
    /// and a stray frame leaks; validator emits "unexpected child element
    /// 'invertIfNegative'").
    /// </summary>
    internal static void InsertSeriesChildInOrder(OpenXmlCompositeElement ser, OpenXmlElement child)
    {
        string[] insertBeforeNames = child.LocalName switch
        {
            // BUG-DUMP-R33-1: verbatim spPr / dPt injection. CT_*Ser order:
            // idx, order, tx, spPr, invertIfNegative, pictureOptions, dPt*,
            // dLbls, trendline, errBars, cat, val, … spPr precedes every
            // styling/data sibling; dPt sits after spPr/invertIfNeg/pictureOpts
            // and before dLbls and the data tail. Appending either at the end
            // makes Word silently ignore it (and the validator rejects it).
            "spPr" => ["invertIfNegative", "pictureOptions", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "bubble3D", "marker", "shape", "smooth", "extLst"],
            "dPt" => ["dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "bubble3D", "shape", "smooth", "extLst"],
            "invertIfNegative" => ["pictureOptions", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "bubble3D", "shape", "smooth", "extLst"],
            // CT_PieSer / CT_DoughnutSer: idx, order, tx?, spPr?, explosion?, dPt*, dLbls?, cat?, val?
            "explosion" => ["dPt", "dLbls", "cat", "val", "extLst"],
            // CT_LineSer / CT_ScatterSer / CT_RadarSer: ..., spPr?, marker?, dPt*,
            // dLbls?, trendline?, errBars?, cat/xVal?, val/yVal?, smooth?, extLst?.
            // marker must precede every data-bearing tail element.
            "marker" => ["dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "smooth", "extLst"],
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
    /// Insert a child into a 3D chart (CT_Bar3DChart / CT_Line3DChart / CT_Area3DChart)
    /// at the correct schema position. All three share the trailing sequence
    /// ..., gapDepth?, [shape? — bar3D only], axId+, extLst?. PowerPoint silently
    /// renders out-of-order children (e.g. shape appended after axId still shows
    /// the cone/cylinder visually) but the validator emits "unexpected child
    /// element 'shape'/'gapDepth' in bar3DChart".
    /// </summary>
    internal static void InsertBar3DChartChildInOrder(OpenXmlCompositeElement chart3d, OpenXmlElement child)
    {
        // bar3D: barDir, grouping, varyColors?, ser*, dLbls?, gapWidth?, gapDepth?, shape?, axId+, extLst?
        // line3D / area3D: grouping?, varyColors?, ser*, dLbls?, dropLines?, gapDepth?, axId+, extLst?
        string[] insertBeforeNames = child.LocalName switch
        {
            "gapDepth" => ["shape", "axId", "extLst"],
            "shape" => ["axId", "extLst"],
            _ => ["axId", "extLst"]
        };
        foreach (var sibling in chart3d.ChildElements)
        {
            if (insertBeforeNames.Contains(sibling.LocalName))
            {
                chart3d.InsertBefore(child, sibling);
                return;
            }
        }
        chart3d.AppendChild(child);
    }

    /// <summary>
    /// Insert a child into a CT_BubbleChart at the correct schema position.
    /// Schema: varyColors?, ser*, dLbls?, bubble3D?, bubbleScale?, showNegBubbles?, sizeRepresents?, axId+, extLst?.
    /// PowerPoint silently renders out-of-order children, but the validator emits
    /// "unexpected child element 'sizeRepresents'/'showNegBubbles'" when they trail axId.
    /// </summary>
    internal static void InsertBubbleChartChildInOrder(OpenXmlCompositeElement bubble, OpenXmlElement child)
    {
        string[] insertBeforeNames = child.LocalName switch
        {
            "bubble3D" => ["bubbleScale", "showNegBubbles", "sizeRepresents", "axId", "extLst"],
            "bubbleScale" => ["showNegBubbles", "sizeRepresents", "axId", "extLst"],
            "showNegBubbles" => ["sizeRepresents", "axId", "extLst"],
            "sizeRepresents" => ["axId", "extLst"],
            _ => ["axId", "extLst"]
        };
        foreach (var sibling in bubble.ChildElements)
        {
            if (insertBeforeNames.Contains(sibling.LocalName))
            {
                bubble.InsertBefore(child, sibling);
                return;
            }
        }
        bubble.AppendChild(child);
    }

    /// <summary>
    /// Insert a child into a CT_ValAx at the correct schema position.
    /// Tail of CT_ValAx: ..., crossAx, crosses?, crossesAt?, crossBetween?,
    /// majorUnit?, minorUnit?, dispUnits?, extLst?. AppendChild is unsafe when
    /// later siblings already exist (e.g. setting majorUnit after minorUnit
    /// already landed flips the schema order and the OpenXmlValidator rejects
    /// the file with "unexpected child element 'majorUnit'").
    /// </summary>
    internal static void InsertValAxChildInOrder(OpenXmlCompositeElement valAx, OpenXmlElement child)
    {
        string[] insertBeforeNames = child.LocalName switch
        {
            "majorUnit" => ["minorUnit", "dispUnits", "extLst"],
            "minorUnit" => ["dispUnits", "extLst"],
            "dispUnits" => ["extLst"],
            _ => ["extLst"]
        };
        foreach (var sibling in valAx.ChildElements)
        {
            if (insertBeforeNames.Contains(sibling.LocalName))
            {
                valAx.InsertBefore(child, sibling);
                return;
            }
        }
        valAx.AppendChild(child);
    }

    /// <summary>
    /// BUG-DUMP-R34-1: replace (or insert) an axis's <c:spPr> at the schema-
    /// correct position. CT_CatAx / CT_ValAx / CT_DateAx share the same tail
    /// after tickLblPos: spPr?, txPr?, crossAx, crosses?/crossesAt?,
    /// crossBetween?, …, lblAlgn?, lblOffset?, tickLblSkip?, tickMarkSkip?,
    /// noMultiLvlLbl?, extLst?. spPr therefore precedes txPr and every cross*
    /// / label* element. AppendChild lands it after crossAx (always present)
    /// and Word silently ignores the out-of-order spPr — so the axis line never
    /// renders. Any existing spPr (typed C.ChartShapeProperties OR the plain
    /// C.ShapeProperties form the SDK produces after a reload) is removed first
    /// so a verbatim replace is idempotent.
    /// </summary>
    internal static void SetAxisSpPr(OpenXmlCompositeElement axis, OpenXmlElement spPr)
    {
        foreach (var existing in axis.ChildElements
            .Where(e => e.LocalName == "spPr").ToList())
            existing.Remove();
        string[] insertBeforeNames =
        [
            "txPr", "crossAx", "crosses", "crossesAt", "crossBetween",
            "majorUnit", "minorUnit", "dispUnits",
            "auto", "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip",
            "noMultiLvlLbl", "baseTimeUnit", "majorTimeUnit", "minorTimeUnit",
            "extLst"
        ];
        foreach (var sibling in axis.ChildElements)
        {
            if (insertBeforeNames.Contains(sibling.LocalName))
            {
                axis.InsertBefore(spPr, sibling);
                return;
            }
        }
        axis.AppendChild(spPr);
    }

    /// <summary>
    /// BUG-DUMP-R34-1: replace (or insert) the plot-area's <c:spPr> at the
    /// schema-correct position. CT_PlotArea tail: …(chart-group)+, (axis)*,
    /// dTable?, spPr?, extLst?. spPr is therefore the last child before extLst.
    /// Any existing spPr (typed or post-reload plain form) is removed first.
    /// </summary>
    internal static void SetPlotAreaSpPr(OpenXmlCompositeElement plotArea, OpenXmlElement spPr)
    {
        foreach (var existing in plotArea.ChildElements
            .Where(e => e.LocalName == "spPr").ToList())
            existing.Remove();
        var extLst = plotArea.ChildElements.FirstOrDefault(e => e.LocalName == "extLst");
        if (extLst != null) plotArea.InsertBefore(spPr, extLst);
        else plotArea.AppendChild(spPr);
    }

    /// <summary>
    /// BUG-DUMP-R35-1: replace (or insert) an axis's <c:txPr> at the schema-
    /// correct position. CT_CatAx / CT_ValAx / CT_DateAx tail: …, spPr?, txPr?,
    /// crossAx, crosses?/crossesAt?, crossBetween?, …, extLst?. txPr therefore
    /// follows spPr and precedes crossAx and every cross*/label* element.
    /// AppendChild lands it after crossAx and Word silently ignores the
    /// out-of-order txPr — so the per-axis font never applies. Any existing
    /// txPr (typed C.TextProperties OR the plain post-reload form) is removed
    /// first so a verbatim replace is idempotent.
    /// </summary>
    internal static void SetAxisTxPr(OpenXmlCompositeElement axis, OpenXmlElement txPr)
    {
        foreach (var existing in axis.ChildElements
            .Where(e => e.LocalName == "txPr").ToList())
            existing.Remove();
        string[] insertBeforeNames =
        [
            "crossAx", "crosses", "crossesAt", "crossBetween",
            "majorUnit", "minorUnit", "dispUnits",
            "auto", "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip",
            "noMultiLvlLbl", "baseTimeUnit", "majorTimeUnit", "minorTimeUnit",
            "extLst"
        ];
        foreach (var sibling in axis.ChildElements)
        {
            if (insertBeforeNames.Contains(sibling.LocalName))
            {
                axis.InsertBefore(txPr, sibling);
                return;
            }
        }
        axis.AppendChild(txPr);
    }

    /// <summary>
    /// Insert a child into the CT_Chart element at the correct schema position.
    /// Schema: title?, autoTitleDeleted?, pivotFmts?, view3D?, floor?, sideWall?,
    /// backWall?, plotArea, legend?, plotVisOnly?, dispBlanksAs?, showDLblsOverMax?, extLst?.
    /// AppendChild leaves trailing elements (plotVisOnly, dispBlanksAs) in the wrong
    /// order when applied after siblings already exist; PowerPoint silently honors
    /// the value, but OpenXmlValidator rejects with 'unexpected child element'.
    /// </summary>
    internal static void InsertChartChildInOrder(OpenXmlCompositeElement chart, OpenXmlElement child)
    {
        string[] insertBeforeNames = child.LocalName switch
        {
            "plotVisOnly" => ["dispBlanksAs", "showDLblsOverMax", "extLst"],
            "dispBlanksAs" => ["showDLblsOverMax", "extLst"],
            "showDLblsOverMax" => ["extLst"],
            _ => ["extLst"]
        };
        foreach (var sibling in chart.ChildElements)
        {
            if (insertBeforeNames.Contains(sibling.LocalName))
            {
                chart.InsertBefore(child, sibling);
                return;
            }
        }
        chart.AppendChild(child);
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
