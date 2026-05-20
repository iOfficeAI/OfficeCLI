// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

internal static partial class ChartHelper
{
    // ==================== Axis by @role path routing ====================
    //
    // Surfaces /chart[N]/axis[@role=ROLE] where ROLE ∈ {category, value, value2, series}.
    // Per schemas/help/pptx/chart-axis.json. Shared across Pptx / Word / Excel handlers.

    /// <summary>
    /// Locate the C.* axis element in the plot area corresponding to the given role.
    /// Returns null if not present.
    /// </summary>
    private static OpenXmlElement? FindAxisByRole(C.PlotArea plotArea, string role)
    {
        switch (role.ToLowerInvariant())
        {
            case "category":
                return (OpenXmlElement?)plotArea.Elements<C.CategoryAxis>().FirstOrDefault()
                    ?? plotArea.Elements<C.DateAxis>().FirstOrDefault();
            case "value":
                return plotArea.Elements<C.ValueAxis>().FirstOrDefault();
            case "value2":
                return plotArea.Elements<C.ValueAxis>().Skip(1).FirstOrDefault();
            case "series":
                return plotArea.Elements<C.SeriesAxis>().FirstOrDefault();
            default:
                return null;
        }
    }

    /// <summary>
    /// Build a DocumentNode describing the axis identified by <paramref name="role"/>.
    /// Returns null if the chart has no plot area or no matching axis.
    /// </summary>
    internal static DocumentNode? BuildAxisNode(C.ChartSpace chartSpace, string role, string path)
    {
        var chart = chartSpace?.GetFirstChild<C.Chart>();
        var plotArea = chart?.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return null;

        var axis = FindAxisByRole(plotArea, role);
        if (axis == null) return null;

        var node = new DocumentNode { Path = path, Type = "axis" };
        node.Format["role"] = role.ToLowerInvariant();

        // Title (axis own title, not chart title)
        var axisTitle = axis.GetFirstChild<C.Title>();
        var axisTitleText = axisTitle?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (axisTitleText != null) node.Format["title"] = axisTitleText;

        // Visible: true unless C.Delete is set truthy
        var deleteEl = axis.GetFirstChild<C.Delete>();
        var deleted = deleteEl?.Val?.Value == true;
        node.Format["visible"] = (!deleted).ToString().ToLowerInvariant();

        // Scaling min/max — only meaningful on value axes
        if (role.Equals("value", StringComparison.OrdinalIgnoreCase)
            || role.Equals("value2", StringComparison.OrdinalIgnoreCase))
        {
            var scaling = axis.GetFirstChild<C.Scaling>();
            var minEl = scaling?.GetFirstChild<C.MinAxisValue>();
            var maxEl = scaling?.GetFirstChild<C.MaxAxisValue>();
            if (minEl?.Val?.HasValue == true)
                node.Format["min"] = minEl.Val.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (maxEl?.Val?.HasValue == true)
                node.Format["max"] = maxEl.Val.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var logBaseEl = scaling?.GetFirstChild<C.LogBase>();
            if (logBaseEl?.Val?.HasValue == true)
                node.Format["logBase"] = logBaseEl.Val.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);

            // MajorUnit/MinorUnit — value axis tick intervals (axis-level reader; mirrors Setter mutation)
            var majorUnitEl = axis.GetFirstChild<C.MajorUnit>();
            if (majorUnitEl?.Val?.HasValue == true)
                node.Format["majorUnit"] = majorUnitEl.Val.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var minorUnitEl = axis.GetFirstChild<C.MinorUnit>();
            if (minorUnitEl?.Val?.HasValue == true)
                node.Format["minorUnit"] = minorUnitEl.Val.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);

            // DisplayUnits — value axis label units (axis-level reader; chart-level Reader emits same key)
            var dispUnitsEl = axis.GetFirstChild<C.DisplayUnits>();
            var builtInUnit = dispUnitsEl?.GetFirstChild<C.BuiltInUnit>()?.Val;
            if (builtInUnit?.HasValue == true)
                node.Format["dispUnits"] = builtInUnit.InnerText;
        }

        // NumberingFormat — applies to any axis role per schema (chart-axis.json `format`)
        var numFmt = axis.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value;
        if (numFmt != null && numFmt != "General") node.Format["format"] = numFmt;

        // Gridline presence
        node.Format["majorGridlines"] = (axis.GetFirstChild<C.MajorGridlines>() != null)
            .ToString().ToLowerInvariant();
        node.Format["minorGridlines"] = (axis.GetFirstChild<C.MinorGridlines>() != null)
            .ToString().ToLowerInvariant();

        // Axis orientation (value/category — schema applies to both via scaling)
        var scalingForOrient = axis.GetFirstChild<C.Scaling>();
        var axisOrient = scalingForOrient?.GetFirstChild<C.Orientation>()?.Val;
        if (axisOrient?.HasValue == true && axisOrient.InnerText == "maxMin")
            node.Format["axisOrientation"] = "maxMin";

        // Tick marks — mirror chart-level reader (R43-1)
        var majorTick = axis.GetFirstChild<C.MajorTickMark>()?.Val;
        if (majorTick?.HasValue == true) node.Format["majorTickMark"] = majorTick.InnerText;
        var minorTick = axis.GetFirstChild<C.MinorTickMark>()?.Val;
        if (minorTick?.HasValue == true) node.Format["minorTickMark"] = minorTick.InnerText;

        // Tick label position
        var tickLblPos = axis.GetFirstChild<C.TickLabelPosition>()?.Val;
        if (tickLblPos?.HasValue == true) node.Format["tickLabelPos"] = tickLblPos.InnerText;

        // Crossing (value axis vocabulary; on category axis these are inert) — R43-2
        if (axis is OpenXmlCompositeElement axCross)
        {
            var crossesVal = axCross.GetFirstChild<C.Crosses>()?.Val;
            if (crossesVal?.HasValue == true) node.Format["crosses"] = crossesVal.InnerText;
            var crossesAtVal = axCross.GetFirstChild<C.CrossesAt>()?.Val?.Value;
            if (crossesAtVal != null) node.Format["crossesAt"] = crossesAtVal;
            var crossBetween = axCross.GetFirstChild<C.CrossBetween>()?.Val;
            if (crossBetween?.HasValue == true) node.Format["crossBetween"] = crossBetween.InnerText;
        }

        // Category-axis specifics — labelOffset, tickLabelSkip
        if (role.Equals("category", StringComparison.OrdinalIgnoreCase))
        {
            var labelOffsetVal = axis.GetFirstChild<C.LabelOffset>()?.Val?.Value;
            if (labelOffsetVal != null && labelOffsetVal != 100)
                node.Format["labelOffset"] = labelOffsetVal;
            var tickLblSkipVal = axis.GetFirstChild<C.TickLabelSkip>()?.Val?.Value;
            if (tickLblSkipVal != null && tickLblSkipVal > 1)
                node.Format["tickLabelSkip"] = tickLblSkipVal;
        }

        // Label rotation from TextProperties BodyProperties.Rotation (60000 per degree)
        var txPr = axis.GetFirstChild<C.TextProperties>();
        var bodyPr = txPr?.GetFirstChild<Drawing.BodyProperties>();
        if (bodyPr?.Rotation?.HasValue == true)
        {
            var deg = bodyPr.Rotation.Value / 60000.0;
            node.Format["labelRotation"] = deg.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        }

        return node;
    }

    /// <summary>
    /// Translate role-scoped Set properties into the existing dotted-key vocabulary
    /// consumed by <see cref="SetChartProperties(ChartPart, Dictionary{string, string})"/>
    /// and forward the call. Returns the list of unsupported keys.
    /// </summary>
    internal static List<string> SetAxisProperties(
        ChartPart chartPart, string role, Dictionary<string, string> properties)
    {
        var normalizedRole = role.ToLowerInvariant();
        var translated = new Dictionary<string, string>();
        var directlyHandled = new List<string>();

        // Resolve target axis once for direct-apply paths.
        var chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
        var plotArea = chart?.GetFirstChild<C.PlotArea>();
        var targetAxis = plotArea != null ? FindAxisByRole(plotArea, normalizedRole) : null;

        foreach (var (key, value) in properties)
        {
            var lower = key.ToLowerInvariant();
            switch (lower)
            {
                case "title":
                    // Map role → existing axis-title keys already handled by SetChartProperties.
                    // category/series → cattitle; value/value2 → axistitle.
                    if (normalizedRole is "category" or "series")
                        translated["cattitle"] = value;
                    else
                        translated["axistitle"] = value;
                    break;

                case "min":
                    // CONSISTENCY(chart/axis-role-write): the legacy `axismin` key
                    // always targets the primary value axis. For role=value2 we must
                    // write to the secondary axis directly to mirror BuildAxisNode's
                    // Skip(1) read path. Same for max/crosses/crossesat below.
                    if (normalizedRole == "value2" && targetAxis is OpenXmlCompositeElement minAx2)
                    {
                        var scaling = minAx2.GetFirstChild<C.Scaling>();
                        if (scaling != null)
                        {
                            scaling.RemoveAllChildren<C.MinAxisValue>();
                            scaling.AppendChild(new C.MinAxisValue { Val = ParseHelpers.SafeParseDouble(value, "min") });
                        }
                        directlyHandled.Add(key);
                    }
                    else
                    {
                        translated["axismin"] = value;
                    }
                    break;

                case "max":
                    if (normalizedRole == "value2" && targetAxis is OpenXmlCompositeElement maxAx2)
                    {
                        var scaling = maxAx2.GetFirstChild<C.Scaling>();
                        if (scaling != null)
                        {
                            scaling.RemoveAllChildren<C.MaxAxisValue>();
                            var maxEl = new C.MaxAxisValue { Val = ParseHelpers.SafeParseDouble(value, "max") };
                            // Schema order: logBase?, orientation, max?, min? — insert max after orientation
                            var orient = scaling.GetFirstChild<C.Orientation>();
                            if (orient != null) orient.InsertAfterSelf(maxEl);
                            else scaling.PrependChild(maxEl);
                        }
                        directlyHandled.Add(key);
                    }
                    else
                    {
                        translated["axismax"] = value;
                    }
                    break;

                case "crosses":
                    if (normalizedRole == "value2" && targetAxis is OpenXmlCompositeElement crsAx2)
                    {
                        crsAx2.RemoveAllChildren<C.Crosses>();
                        crsAx2.RemoveAllChildren<C.CrossesAt>();
                        var crossVal = value.ToLowerInvariant() switch
                        {
                            "max" => C.CrossesValues.Maximum,
                            "min" => C.CrossesValues.Minimum,
                            _ => C.CrossesValues.AutoZero
                        };
                        var newCrosses = new C.Crosses { Val = crossVal };
                        var cbBefore = crsAx2.GetFirstChild<C.CrossBetween>();
                        if (cbBefore != null) crsAx2.InsertBefore(newCrosses, cbBefore);
                        else crsAx2.AppendChild(newCrosses);
                        directlyHandled.Add(key);
                    }
                    else
                    {
                        translated["crosses"] = value;
                    }
                    break;

                case "crossesat":
                    if (normalizedRole == "value2" && targetAxis is OpenXmlCompositeElement crsAtAx2)
                    {
                        crsAtAx2.RemoveAllChildren<C.Crosses>();
                        crsAtAx2.RemoveAllChildren<C.CrossesAt>();
                        var newCrossesAt = new C.CrossesAt { Val = ParseHelpers.SafeParseDouble(value, "crossesAt") };
                        var cbBefore2 = crsAtAx2.GetFirstChild<C.CrossBetween>();
                        if (cbBefore2 != null) crsAtAx2.InsertBefore(newCrossesAt, cbBefore2);
                        else crsAtAx2.AppendChild(newCrossesAt);
                        directlyHandled.Add(key);
                    }
                    else
                    {
                        translated["crossesat"] = value;
                    }
                    break;

                case "labelrotation":
                    // Existing setter already understands xaxis.labelrotation / yaxis.labelrotation.
                    translated[normalizedRole is "category" or "series"
                        ? "xaxis.labelrotation"
                        : "yaxis.labelrotation"] = value;
                    break;

                case "visible":
                    // Map by role to the existing role-specific cataxisvisible/valaxisvisible
                    // keys. value/value2/series are not split in the legacy setter, so for
                    // value2 we apply directly on the resolved axis.
                    if (normalizedRole is "category")
                        translated["cataxisvisible"] = value;
                    else if (normalizedRole is "value" or "series")
                        translated["valaxisvisible"] = value;
                    else if (targetAxis is OpenXmlCompositeElement axCe)
                    {
                        axCe.RemoveAllChildren<C.Delete>();
                        axCe.InsertAfter(
                            new C.Delete { Val = !ParseHelpers.IsTruthy(value) },
                            axCe.GetFirstChild<C.Scaling>());
                        directlyHandled.Add(key);
                    }
                    else
                    {
                        directlyHandled.Add(key); // axis missing; treat as no-op silently
                    }
                    break;

                case "majortickmark":
                case "minortickmark":
                case "majortick":
                case "minortick":
                {
                    // CONSISTENCY(chart/axis-role-write): legacy SetChartProperties
                    // applies tickmark to every ValueAxis and CategoryAxis. Under a
                    // role-scoped write we must only touch the resolved axis.
                    if (targetAxis is OpenXmlCompositeElement axTick)
                    {
                        var tickVal = ParseTickMark(value);
                        if (lower == "majortickmark" || lower == "majortick")
                        {
                            axTick.RemoveAllChildren<C.MajorTickMark>();
                            InsertAxisChildInOrder(axTick, new C.MajorTickMark { Val = tickVal });
                        }
                        else
                        {
                            axTick.RemoveAllChildren<C.MinorTickMark>();
                            InsertAxisChildInOrder(axTick, new C.MinorTickMark { Val = tickVal });
                        }
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "logbase":
                {
                    // Schema: logBase only valid on role=value/value2; category/series → ignore.
                    if (normalizedRole is not ("value" or "value2"))
                    {
                        directlyHandled.Add(key);
                        break;
                    }
                    if (targetAxis is OpenXmlCompositeElement axLb)
                    {
                        var scaling = axLb.GetFirstChild<C.Scaling>();
                        if (scaling != null)
                        {
                            scaling.RemoveAllChildren<C.LogBase>();
                            if (value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                                value.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                                value.Equals("log", StringComparison.OrdinalIgnoreCase) ||
                                value == "1")
                            {
                                scaling.PrependChild(new C.LogBase { Val = 10d });
                            }
                            else if (!value.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                                     !value.Equals("linear", StringComparison.OrdinalIgnoreCase) &&
                                     !value.Equals("false", StringComparison.OrdinalIgnoreCase) &&
                                     !value.Equals("no", StringComparison.OrdinalIgnoreCase) &&
                                     value != "0")
                            {
                                var logVal = ParseHelpers.SafeParseDouble(value, "logBase");
                                // ST_LogBase: minInclusive=2.0, maxExclusive=1000.0 — reject
                                // fractional and out-of-band values so Excel doesn't silently
                                // ghost-rewrite the chart back to linear.
                                if (logVal < 2.0 || logVal >= 1000.0)
                                    throw new ArgumentException($"Invalid logBase '{value}': must be in the OOXML range [2, 1000) (ST_LogBase).");
                                scaling.PrependChild(new C.LogBase { Val = logVal });
                            }
                        }
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "format":
                {
                    // Number-format string written as the axis's NumberingFormat child.
                    // Schema declares format on all roles; apply directly on the resolved axis.
                    if (targetAxis is OpenXmlCompositeElement axNf)
                    {
                        axNf.RemoveAllChildren<C.NumberingFormat>();
                        var nf = new C.NumberingFormat { FormatCode = value, SourceLinked = false };
                        // Schema order: ...title, numFmt, majorTickMark... — insert before majorTickMark
                        var nfBefore = axNf.GetFirstChild<C.MajorTickMark>();
                        if (nfBefore != null) axNf.InsertBefore(nf, nfBefore);
                        else axNf.AppendChild(nf);
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "ticklabelpos":
                case "ticklabelposition":
                {
                    // CONSISTENCY(chart/axis-role-write): legacy SetChartProperties
                    // tickLabelPos sweeps every ValueAxis + CategoryAxis. Role-scoped
                    // write must only mutate the resolved axis. (R43-4)
                    if (targetAxis is OpenXmlCompositeElement axTlp)
                    {
                        var tlPos = value.ToLowerInvariant() switch
                        {
                            "none" => C.TickLabelPositionValues.None,
                            "high" or "top" => C.TickLabelPositionValues.High,
                            "low" or "bottom" => C.TickLabelPositionValues.Low,
                            _ => C.TickLabelPositionValues.NextTo
                        };
                        axTlp.RemoveAllChildren<C.TickLabelPosition>();
                        InsertAxisChildInOrder(axTlp, new C.TickLabelPosition { Val = tlPos });
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "labeloffset":
                {
                    // Category-axis-only per OOXML schema. Skip on other roles.
                    if (normalizedRole != "category") { directlyHandled.Add(key); break; }
                    if (targetAxis is OpenXmlCompositeElement axLo)
                    {
                        axLo.RemoveAllChildren<C.LabelOffset>();
                        axLo.AppendChild(new C.LabelOffset { Val = (ushort)ParseHelpers.SafeParseInt(value, "labelOffset") });
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "ticklabelskip":
                case "tickskip":
                {
                    if (normalizedRole != "category") { directlyHandled.Add(key); break; }
                    if (targetAxis is OpenXmlCompositeElement axTls)
                    {
                        axTls.RemoveAllChildren<C.TickLabelSkip>();
                        axTls.AppendChild(new C.TickLabelSkip { Val = ParseHelpers.SafeParseInt(value, "tickLabelSkip") });
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "crossbetween":
                {
                    // Schema: crossBetween only valid on value/value2; on category/series ignore.
                    if (normalizedRole is not ("value" or "value2")) { directlyHandled.Add(key); break; }
                    if (targetAxis is OpenXmlCompositeElement axCb)
                    {
                        axCb.RemoveAllChildren<C.CrossBetween>();
                        var cbVal = value.ToLowerInvariant() switch
                        {
                            "midcat" or "midpoint" => C.CrossBetweenValues.MidpointCategory,
                            _ => C.CrossBetweenValues.Between
                        };
                        // CT_ValAx schema: ..., crossAx, crosses?, crossesAt?,
                        // crossBetween?, majorUnit?, minorUnit?, dispUnits?, extLst?.
                        // AppendChild lands it after majorUnit which PowerPoint
                        // rejects ("unexpected child element 'crossBetween'").
                        var cb = new C.CrossBetween { Val = cbVal };
                        var cbAnchor = axCb.GetFirstChild<C.CrossesAt>() as OpenXmlElement
                            ?? axCb.GetFirstChild<C.Crosses>() as OpenXmlElement
                            ?? axCb.GetFirstChild<C.CrossingAxis>() as OpenXmlElement;
                        if (cbAnchor != null) cbAnchor.InsertAfterSelf(cb);
                        else axCb.AppendChild(cb);
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "axisorientation":
                case "orientation":
                case "axisreverse":
                {
                    // Role-scoped orientation write — legacy `axisorientation` in
                    // SetChartProperties writes to the primary value axis only,
                    // ignoring role. Apply directly on the resolved axis. (R43-3)
                    if (targetAxis is OpenXmlCompositeElement axOr)
                    {
                        var scaling = axOr.GetFirstChild<C.Scaling>();
                        if (scaling != null)
                        {
                            scaling.RemoveAllChildren<C.Orientation>();
                            var orientVal = (ParseHelpers.IsValidBooleanString(value) && ParseHelpers.IsTruthy(value)) ||
                                            value.Equals("maxmin", StringComparison.OrdinalIgnoreCase)
                                ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax;
                            scaling.PrependChild(new C.Orientation { Val = orientVal });
                        }
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "majorunit":
                case "minorunit":
                {
                    // Schema: majorUnit / minorUnit only valid on value/value2.
                    // Without this direct-apply branch the role-scoped Set on
                    // role=value2 falls through to the chart-level case which
                    // always grabs the primary ValueAxis, so the secondary
                    // axis silently retained its old tick interval.
                    if (normalizedRole is not ("value" or "value2"))
                    {
                        directlyHandled.Add(key);
                        break;
                    }
                    if (targetAxis is OpenXmlCompositeElement axMu)
                    {
                        var unit = ParseHelpers.SafeParseDouble(value, lower);
                        if (!(unit > 0))
                            throw new ArgumentException(
                                $"Invalid {lower} '{value}': must be a positive number (OOXML ST_AxisUnit > 0).");
                        if (lower == "majorunit")
                        {
                            axMu.RemoveAllChildren<C.MajorUnit>();
                            InsertValAxChildInOrder(axMu, new C.MajorUnit { Val = unit });
                        }
                        else
                        {
                            axMu.RemoveAllChildren<C.MinorUnit>();
                            InsertValAxChildInOrder(axMu, new C.MinorUnit { Val = unit });
                        }
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "majorgridlines":
                case "minorgridlines":
                {
                    if (targetAxis is OpenXmlCompositeElement axCe)
                    {
                        var enable = !value.Equals("none", StringComparison.OrdinalIgnoreCase)
                            && !value.Equals("false", StringComparison.OrdinalIgnoreCase);
                        if (lower == "majorgridlines")
                        {
                            axCe.RemoveAllChildren<C.MajorGridlines>();
                            if (enable)
                            {
                                var gl = new C.MajorGridlines();
                                if (!value.Equals("true", StringComparison.OrdinalIgnoreCase))
                                    gl.AppendChild(BuildLineShapeProperties(value));
                                axCe.InsertAfter(gl, axCe.GetFirstChild<C.AxisPosition>());
                            }
                        }
                        else
                        {
                            axCe.RemoveAllChildren<C.MinorGridlines>();
                            if (enable)
                            {
                                var gl = new C.MinorGridlines();
                                if (!value.Equals("true", StringComparison.OrdinalIgnoreCase))
                                    gl.AppendChild(BuildLineShapeProperties(value));
                                var afterEl = (OpenXmlElement?)axCe.GetFirstChild<C.MajorGridlines>()
                                    ?? axCe.GetFirstChild<C.AxisPosition>();
                                if (afterEl != null) axCe.InsertAfter(gl, afterEl);
                            }
                        }
                    }
                    directlyHandled.Add(key);
                    break;
                }

                default:
                    // Forward unknown keys verbatim; SetChartProperties will flag them as unsupported.
                    translated[key] = value;
                    break;
            }
        }

        var unsupported = translated.Count > 0
            ? SetChartProperties(chartPart, translated)
            : new List<string>();
        // directlyHandled keys are already applied; do not surface as unsupported.
        return unsupported;
    }
}
