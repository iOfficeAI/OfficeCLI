// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
        // CONSISTENCY(axis-title-styling): mirror the Set surface — when callers
        // can write title.font/color/size on the axis Set path, they must also
        // be able to read them back on the axis Get. Pull from the first run's
        // rPr (and fall back to defRPr) so the readback matches what was set.
        if (axisTitle != null)
        {
            var firstAxisTitleRun = axisTitle.Descendants<Drawing.Run>().FirstOrDefault();
            var firstAxisRPr = firstAxisTitleRun?.RunProperties;
            var axisDefRPr = axisTitle.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();

            var atFont = firstAxisRPr?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                ?? axisDefRPr?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
            if (!string.IsNullOrEmpty(atFont)) node.Format["title.font"] = atFont;

            var atSize = firstAxisRPr?.FontSize?.Value ?? axisDefRPr?.FontSize?.Value;
            if (atSize.HasValue)
                node.Format["title.size"] = $"{atSize.Value / 100.0:0.##}pt";

            var atBold = firstAxisRPr?.Bold?.Value ?? axisDefRPr?.Bold?.Value;
            if (atBold == true) node.Format["title.bold"] = "true";

            // Color from rPr's solidFill (or defRPr's). Mirror ParseHelpers
            // canonical "#RRGGBB" used elsewhere in chart readback.
            var atSolid = firstAxisRPr?.GetFirstChild<Drawing.SolidFill>()
                ?? axisDefRPr?.GetFirstChild<Drawing.SolidFill>();
            var atRgbEl = atSolid?.GetFirstChild<Drawing.RgbColorModelHex>();
            if (atRgbEl?.Val?.Value is { } atRgb && !string.IsNullOrEmpty(atRgb))
                node.Format["title.color"] = ParseHelpers.FormatHexColor(atRgb);
        }

        // Visible: true unless C.Delete is set truthy
        var deleteEl = axis.GetFirstChild<C.Delete>();
        var deleted = deleteEl?.Val?.Value == true;
        node.Format["visible"] = (!deleted).ToString().ToLowerInvariant();

        // Date-axis base time unit (days/months/years) — drives the tick
        // spacing of the plotted range; a years-based source replayed as
        // days renders hairline bars.
        if (axis is C.DateAxis dateAxEl)
        {
            var btu = dateAxEl.GetFirstChild<C.BaseTimeUnit>()?.Val;
            if (btu?.HasValue == true)
                node.Format["baseTimeUnit"] = btu.InnerText;
        }

        // Scaling min/max — meaningful on value axes and on a DateAxis
        // (its min/max are date serials that window the plotted range).
        if (role.Equals("value", StringComparison.OrdinalIgnoreCase)
            || role.Equals("value2", StringComparison.OrdinalIgnoreCase)
            || (role.Equals("category", StringComparison.OrdinalIgnoreCase)
                && axis is C.DateAxis))
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

        // Gridline presence + detail. Mirror the chart-level Reader: presence is a
        // boolean, but the gridline's solidFill color / outline width / dash are
        // surfaced via ReadGridlineDetail so a `majorgridlines=COLOR:WIDTH` Set is
        // visible on Get (gridlineColor / gridlineWidth / gridlineDash).
        var majorGridlines = axis.GetFirstChild<C.MajorGridlines>();
        node.Format["majorGridlines"] = (majorGridlines != null).ToString().ToLowerInvariant();
        if (majorGridlines != null) ReadGridlineDetail(majorGridlines, node, "gridline");
        var minorGridlines = axis.GetFirstChild<C.MinorGridlines>();
        node.Format["minorGridlines"] = (minorGridlines != null).ToString().ToLowerInvariant();
        if (minorGridlines != null) ReadGridlineDetail(minorGridlines, node, "minorGridline");

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
        var pendingAxisTitleStyling = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);

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
                case "axistitle":
                case "vtitle":
                    // Map role → existing axis-title keys already handled by SetChartProperties.
                    // category/series → cattitle; value → axistitle (primary value axis).
                    // The common alias `axisTitle` (and `vtitle`) must route here
                    // too — otherwise it falls through to the default and always
                    // targets the PRIMARY value axis, clobbering it for role=value2.
                    if (normalizedRole is "category" or "series")
                        translated["cattitle"] = value;
                    // CONSISTENCY(chart/axis-role-write): the legacy `axistitle`
                    // key always targets the PRIMARY value axis. For role=value2
                    // that would overwrite the primary axis's title and leave the
                    // secondary untouched — write directly to the resolved
                    // secondary axis instead (mirrors min/max/crosses below).
                    else if (normalizedRole == "value2" && targetAxis is OpenXmlCompositeElement titleAx2)
                    {
                        titleAx2.RemoveAllChildren<C.Title>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            ParseHelpers.ValidateXmlText(value, "axisTitle");
                            var insertAfter = (OpenXmlElement?)titleAx2.GetFirstChild<C.MinorGridlines>()
                                ?? (OpenXmlElement?)titleAx2.GetFirstChild<C.MajorGridlines>()
                                ?? titleAx2.GetFirstChild<C.AxisPosition>();
                            var newTitle = BuildChartTitle(value);
                            if (insertAfter != null) titleAx2.InsertAfter(newTitle, insertAfter);
                            else titleAx2.AppendChild(newTitle);
                        }
                        directlyHandled.Add(key);
                    }
                    else
                        translated["axistitle"] = value;
                    break;

                case "min":
                case "axismin":
                    // CONSISTENCY(chart/axis-role-write): the legacy `axismin` key
                    // always targets the primary value axis. For role=value2 we must
                    // write to the secondary axis directly to mirror BuildAxisNode's
                    // Skip(1) read path. Same for max/crosses/crossesat below.
                    // Category (date) and series axes must also write directly:
                    // the legacy `axismin` fallback targets the primary VALUE
                    // axis, which would clobber the wrong scaling.
                    // `axismin` is the spelling the chart schema declares as an
                    // alias of `min` (schemas/help/_shared/chart.json), so it has
                    // to take this branch too: while it fell to `default:` the
                    // forwarded legacy key wrote the PRIMARY value axis even
                    // though the path addressed the secondary one (issue #503).
                    // `axismax` / `axisnumfmt` below carry their own aliases for
                    // the same reason.
                    // CONSISTENCY(chart/axis-alias-path): an alias and its
                    // canonical name must resolve identically on a role path.
                    if (normalizedRole is "value2" or "category" or "series"
                        && targetAxis is OpenXmlCompositeElement minAx2)
                    {
                        var scaling = minAx2.GetFirstChild<C.Scaling>();
                        if (scaling != null)
                        {
                            var minV = ParseHelpers.SafeParseDouble(value, "min");
                            // A log-scaled axis cannot have min <= 0 (Excel
                            // refuses the file, 0x800A03EC).
                            if (minV <= 0 && scaling.GetFirstChild<C.LogBase>() != null)
                                throw new ArgumentException(
                                    $"min={value} is invalid on a log-scaled axis: a logarithmic axis minimum must be greater than 0.");
                            scaling.RemoveAllChildren<C.MinAxisValue>();
                            // CT_Scaling order: logBase, orientation, max, min —
                            // min is last, so append is always valid.
                            scaling.AppendChild(new C.MinAxisValue { Val = minV });
                        }
                        directlyHandled.Add(key);
                    }
                    else
                    {
                        translated["axismin"] = value;
                    }
                    break;

                case "max":
                case "axismax":
                    if (normalizedRole is "value2" or "category" or "series"
                        && targetAxis is OpenXmlCompositeElement maxAx2)
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
                        // Same-type only — see ChartHelper.Setter.cs case "crosses"
                        // for the mutual-remove bug rationale.
                        // Validate BEFORE mutating (atomicity).
                        var crossVal = value.ToLowerInvariant() switch
                        {
                            "max" => C.CrossesValues.Maximum,
                            "min" => C.CrossesValues.Minimum,
                            "autozero" => C.CrossesValues.AutoZero,
                            _ => throw new ArgumentException($"Invalid 'crosses' value: '{value}'. Valid: autoZero, max, min.")
                        };
                        crsAx2.RemoveAllChildren<C.Crosses>();
                        var newCrosses = new C.Crosses { Val = crossVal };
                        var crsAnchor = crsAx2.GetFirstChild<C.CrossesAt>() as OpenXmlElement
                            ?? crsAx2.GetFirstChild<C.CrossBetween>() as OpenXmlElement;
                        if (crsAnchor != null) crsAx2.InsertBefore(newCrosses, crsAnchor);
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
                        // Same-type only.
                        var crossesAtVal2 = ParseHelpers.SafeParseDouble(value, "crossesAt");
                        crsAtAx2.RemoveAllChildren<C.CrossesAt>();
                        var newCrossesAt = new C.CrossesAt { Val = crossesAtVal2 };
                        var cbBefore2 = crsAtAx2.GetFirstChild<C.CrossBetween>();
                        if (cbBefore2 != null) crsAtAx2.InsertBefore(newCrossesAt, cbBefore2);
                        else crsAtAx2.AppendChild(newCrossesAt);
                        // CONSISTENCY(chart/crossesat-overrides-crosses): mirror
                        // ChartHelper.Setter.cs case "crossesat" — suppress the
                        // default <c:crosses val="autoZero"/> seeded by the axis
                        // builder when the caller did not request `crosses`
                        // explicitly, so dump→replay does not surface a
                        // spurious crosses=autoZero (R57 tester-1).
                        if (!properties.ContainsKey("crosses"))
                            crsAtAx2.RemoveAllChildren<C.Crosses>();
                        directlyHandled.Add(key);
                    }
                    else
                    {
                        translated["crossesat"] = value;
                    }
                    break;

                case "labelrotation":
                    // CONSISTENCY(chart/axis-role-write): the legacy
                    // xaxis.labelrotation / yaxis.labelrotation keys target the
                    // CategoryAxis(es) / primary ValueAxis respectively, ignoring
                    // role. For value2 (secondary value axis) the legacy
                    // yaxis.labelrotation stamps EVERY ValueAxis, corrupting the
                    // primary; for series it would route to xaxis.labelrotation and
                    // mutate the CategoryAxis instead of the SeriesAxis. Apply
                    // directly on the resolved target axis to mirror BuildAxisNode's
                    // per-axis labelRotation read.
                    if (normalizedRole is "value2" or "series"
                        && targetAxis is OpenXmlCompositeElement axRot)
                    {
                        if (double.TryParse(value, System.Globalization.NumberStyles.Float,
                                System.Globalization.CultureInfo.InvariantCulture, out var rotDeg))
                        {
                            var rotAttrVal = ((int)(rotDeg * 60000))
                                .ToString(System.Globalization.CultureInfo.InvariantCulture);
                            ApplyAxisLabelRotation(axRot, rotAttrVal);
                            directlyHandled.Add(key);
                        }
                        else
                        {
                            translated[key] = value; // let SetChartProperties flag the bad value
                        }
                    }
                    else
                    {
                        // Existing setter already understands xaxis.labelrotation / yaxis.labelrotation.
                        translated[normalizedRole is "category"
                            ? "xaxis.labelrotation"
                            : "yaxis.labelrotation"] = value;
                    }
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
                            // Resolve+validate BEFORE mutating so a bad numeric
                            // base doesn't wipe the prior valid log scale.
                            double? newLogBase;
                            if (value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                                value.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                                value.Equals("log", StringComparison.OrdinalIgnoreCase))
                            {
                                // "1" was historically truthy shorthand here too;
                                // routed through SafeParseDouble + range-check below
                                // so logBase=1 surfaces as ArgumentException.
                                newLogBase = 10d;
                            }
                            else if (value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
                                     value.Equals("linear", StringComparison.OrdinalIgnoreCase) ||
                                     value.Equals("false", StringComparison.OrdinalIgnoreCase) ||
                                     value.Equals("no", StringComparison.OrdinalIgnoreCase))
                            {
                                newLogBase = null; // remove log scale (linear)
                            }
                            else
                            // "0" dropped as a falsy synonym for the same
                            // reason as the Setter.cs site — falls into the
                            // range check below and throws.
                            {
                                var logVal = ParseHelpers.SafeParseDouble(value, "logBase");
                                // ST_LogBase: minInclusive=2.0, maxInclusive=1000.0 — reject
                                // out-of-band values so Excel doesn't silently
                                // ghost-rewrite the chart back to linear.
                                if (logVal < 2.0 || logVal > 1000.0)
                                    throw new ArgumentException($"Invalid logBase '{value}': must be in the OOXML range [2, 1000] (ST_LogBase).");
                                newLogBase = logVal;
                            }
                            // A log scale requires axis min > 0 (Excel refuses
                            // the file, 0x800A03EC).
                            if (newLogBase != null
                                && scaling.GetFirstChild<C.MinAxisValue>()?.Val?.Value is { } curMin && curMin <= 0)
                                throw new ArgumentException(
                                    $"logBase cannot be enabled while the axis minimum ({curMin}) is <= 0: a logarithmic axis minimum must be greater than 0.");
                            scaling.RemoveAllChildren<C.LogBase>();
                            if (newLogBase != null)
                                scaling.PrependChild(new C.LogBase { Val = newLogBase.Value });
                        }
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "format":
                case "axisnumfmt":
                case "axisnumberformat":
                {
                    // Number-format string written as the axis's NumberingFormat child.
                    // Schema declares format on all roles; apply directly on the resolved axis.
                    // `axisnumfmt` / `axisnumberformat` are its declared alias
                    // spellings — see the `min` case above.
                    if (targetAxis is OpenXmlCompositeElement axNf)
                    {
                        axNf.RemoveAllChildren<C.NumberingFormat>();
                        var nf = new C.NumberingFormat { FormatCode = value, SourceLinked = false };
                        // Schema order: ...title, numFmt, majorTickMark... — insert before majorTickMark
                        var nfBefore = axNf.GetFirstChild<C.MajorTickMark>();
                        if (nfBefore != null) axNf.InsertBefore(nf, nfBefore);
                        else axNf.AppendChild(nf);

                        // Date axis: real PowerPoint only engages date-axis
                        // layout when the CATEGORY numCache carries a
                        // date-looking formatCode — with "General" it renders
                        // an empty plot. Propagate the axis format into every
                        // series' cat numCache/numLit formatCode.
                        if (axNf is C.DateAxis && plotArea != null)
                        {
                            foreach (var catData in plotArea
                                .Descendants<C.CategoryAxisData>())
                            {
                                var cache = (OpenXmlCompositeElement?)catData
                                        .GetFirstChild<C.NumberReference>()
                                        ?.GetFirstChild<C.NumberingCache>()
                                    ?? catData.GetFirstChild<C.NumberLiteral>();
                                var fc = cache?.GetFirstChild<C.FormatCode>();
                                if (fc != null) fc.Text = value;
                            }
                        }
                    }
                    directlyHandled.Add(key);
                    break;
                }

                case "basetimeunit":
                {
                    // Date-axis base time unit round-trip (days/months/years).
                    if (targetAxis is C.DateAxis daxBtu)
                    {
                        var btuVal = value.Trim().ToLowerInvariant() switch
                        {
                            "days" or "day" => (C.TimeUnitValues?)C.TimeUnitValues.Days,
                            "months" or "month" => C.TimeUnitValues.Months,
                            "years" or "year" => C.TimeUnitValues.Years,
                            _ => null,
                        };
                        if (btuVal == null)
                            throw new ArgumentException($"Invalid baseTimeUnit '{value}': expected days, months, or years.");
                        daxBtu.RemoveAllChildren<C.BaseTimeUnit>();
                        var btuEl = new C.BaseTimeUnit { Val = btuVal.Value };
                        // CT_DateAx: …auto?, lblOffset?, baseTimeUnit?, majorUnit?…
                        var btuBefore = (OpenXmlElement?)daxBtu.GetFirstChild<C.MajorUnit>();
                        if (btuBefore != null) daxBtu.InsertBefore(btuEl, btuBefore);
                        else daxBtu.AppendChild(btuEl);
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
                        var tlsVal = ParseHelpers.SafeParseInt(value, "tickLabelSkip");
                        if (tlsVal < 1 || tlsVal > 65535)
                            throw new ArgumentException($"Invalid 'tickLabelSkip' value: '{value}'. Must be an integer 1..65535 (OOXML ST_Skip).");
                        axTls.RemoveAllChildren<C.TickLabelSkip>();
                        axTls.AppendChild(new C.TickLabelSkip { Val = tlsVal });
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

                case "title.font" or "titlefont":
                case "title.size" or "titlesize":
                case "title.color" or "titlecolor":
                case "title.bold" or "titlebold":
                    // CONSISTENCY(axis-title-styling): these used to fall to default
                    // and forward to the chart-level title handler, which then
                    // mutated the wrong title (or returned UNSUPPORTED when no
                    // chart title existed). Buffer them here and apply AFTER the
                    // translated forward — that's where `axisTitle=…` / `title=…`
                    // creates the C.Title on this axis, so we need to operate on
                    // the post-forward DOM.
                    pendingAxisTitleStyling[lower] = value;
                    directlyHandled.Add(key);
                    break;

                default:
                    // Forward unknown keys verbatim; SetChartProperties will flag them as unsupported.
                    translated[key] = value;
                    break;
            }
        }

        var unsupported = translated.Count > 0
            ? SetChartProperties(chartPart, translated)
            : new List<string>();

        // Apply axis-scoped title styling AFTER the chart-level setter has
        // had a chance to create/replace the C.Title (axistitle / cattitle
        // build a fresh title from scratch). Re-resolve targetAxis since the
        // forward may have mutated the plotArea.
        if (pendingAxisTitleStyling.Count > 0)
        {
            var axisAfter = plotArea != null ? FindAxisByRole(plotArea, normalizedRole) : null;
            var axisTitle = (axisAfter as OpenXmlCompositeElement)?.GetFirstChild<C.Title>();

            // R52 bt-2: when title.size / title.bold / title.font / title.color
            // is set on an axis whose title element is missing, auto-create an
            // empty <c:title> block on the resolved axis. Previously these keys
            // were silently shoveled into the unsupported list — asymmetric with
            // the Get side (which surfaces the same keys via DefaultRunProperties
            // on a present-but-empty title) and indistinguishable from a true
            // unknown-key reject. Mirrors BuildChartTitle's empty-title shape.
            if (axisTitle == null && axisAfter is OpenXmlCompositeElement axHost)
            {
                axisTitle = new C.Title(
                    new C.ChartText(
                        new C.RichText(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph(
                                new Drawing.ParagraphProperties(
                                    new Drawing.DefaultRunProperties()),
                                new Drawing.Run(
                                    new Drawing.RunProperties { Language = "en-US" },
                                    new Drawing.Text(string.Empty))))),
                    new C.Overlay { Val = false });
                // Schema order: ...title, numFmt, majorTickMark... — title sits
                // after MinorGridlines / MajorGridlines / AxisPosition. Match
                // the insertion site used by axistitle/cattitle in
                // SetChartProperties so PowerPoint accepts the result.
                var insertAfter = (OpenXmlElement?)axHost.GetFirstChild<C.MinorGridlines>()
                    ?? (OpenXmlElement?)axHost.GetFirstChild<C.MajorGridlines>()
                    ?? axHost.GetFirstChild<C.AxisPosition>();
                if (insertAfter != null) axHost.InsertAfter(axisTitle, insertAfter);
                else axHost.AppendChild(axisTitle);
            }

            foreach (var (axKey, axVal) in pendingAxisTitleStyling)
            {
                if (axisTitle == null) { unsupported.Add(axKey); continue; }
                var norm = axKey.Replace("title.", "").Replace("title", "");

                // R42-B2: title.font accepts either a bare font name ("Arial")
                // or a composite "size:color:fontname" spec ("14:4472C4:Arial").
                // The legacy path stored the entire composite as the LatinFont
                // typeface, producing an invalid font with literal text
                // "14:4472C4:Arial". Detect a `:`-delimited composite and route
                // through BuildDefaultRunPropertiesFromCompoundSpec — identical
                // parsing as the axisfont knob — to fan out size/color/font.
                if (norm == "font" && axVal.Contains(':'))
                {
                    var spec = BuildDefaultRunPropertiesFromCompoundSpec(axVal);
                    foreach (var run in axisTitle.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        // Font size: lift hundredths-of-a-point from the parsed defRp.
                        if (spec.FontSize?.HasValue == true)
                            rPr.FontSize = spec.FontSize.Value;
                        // Color: replace any existing SolidFill with the parsed one.
                        var fill = spec.GetFirstChild<Drawing.SolidFill>();
                        if (fill != null)
                        {
                            rPr.RemoveAllChildren<Drawing.SolidFill>();
                            DrawingEffectsHelper.InsertFillInRunProperties(rPr,
                                (Drawing.SolidFill)fill.CloneNode(true));
                        }
                        // Typeface: replace LatinFont + EastAsianFont.
                        var latin = spec.GetFirstChild<Drawing.LatinFont>();
                        if (latin != null)
                        {
                            rPr.RemoveAllChildren<Drawing.LatinFont>();
                            rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                            rPr.AppendChild((Drawing.LatinFont)latin.CloneNode(true));
                            var ea = spec.GetFirstChild<Drawing.EastAsianFont>();
                            if (ea != null)
                                rPr.AppendChild((Drawing.EastAsianFont)ea.CloneNode(true));
                        }
                    }
                    // Mirror onto defRPr for fallback rendering.
                    var defRpComposite = axisTitle.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
                    if (defRpComposite != null)
                    {
                        if (spec.FontSize?.HasValue == true)
                            defRpComposite.FontSize = spec.FontSize.Value;
                    }
                    continue;
                }

                foreach (var run in axisTitle.Descendants<Drawing.Run>())
                {
                    var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                    switch (norm)
                    {
                        case "font":
                            rPr.RemoveAllChildren<Drawing.LatinFont>();
                            rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                            rPr.AppendChild(new Drawing.LatinFont { Typeface = axVal });
                            rPr.AppendChild(new Drawing.EastAsianFont { Typeface = axVal });
                            break;
                        case "size":
                            var sizeStr = axVal.EndsWith("pt", System.StringComparison.OrdinalIgnoreCase)
                                ? axVal[..^2] : axVal;
                            rPr.FontSize = (int)System.Math.Round(
                                ParseHelpers.SafeParseDouble(sizeStr, "title.size") * 100);
                            break;
                        case "color":
                        {
                            rPr.RemoveAllChildren<Drawing.SolidFill>();
                            var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(axVal);
                            DrawingEffectsHelper.InsertFillInRunProperties(rPr,
                                new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = rgb }));
                            break;
                        }
                        case "bold":
                            rPr.Bold = ParseHelpers.IsTruthy(axVal);
                            break;
                    }
                }
                // Mirror chart-Setter behavior — keep defRPr in sync for size/bold.
                var defRp = axisTitle.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
                if (defRp != null)
                {
                    switch (norm)
                    {
                        case "size":
                            var sizeStr = axVal.EndsWith("pt", System.StringComparison.OrdinalIgnoreCase)
                                ? axVal[..^2] : axVal;
                            defRp.FontSize = (int)System.Math.Round(
                                ParseHelpers.SafeParseDouble(sizeStr, "title.size") * 100);
                            break;
                        case "bold":
                            defRp.Bold = ParseHelpers.IsTruthy(axVal);
                            break;
                    }
                }
            }
        }
        // directlyHandled keys are already applied; do not surface as unsupported.
        return unsupported;
    }
}
