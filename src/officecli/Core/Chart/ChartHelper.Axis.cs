// Copyright 2025 OfficeCli (officecli.ai)
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
        }

        // Gridline presence
        node.Format["majorGridlines"] = (axis.GetFirstChild<C.MajorGridlines>() != null)
            .ToString().ToLowerInvariant();
        node.Format["minorGridlines"] = (axis.GetFirstChild<C.MinorGridlines>() != null)
            .ToString().ToLowerInvariant();

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
                    translated["axismin"] = value;
                    break;

                case "max":
                    translated["axismax"] = value;
                    break;

                case "labelrotation":
                    // Existing setter already understands xaxis.labelrotation / yaxis.labelrotation.
                    translated[normalizedRole is "category" or "series"
                        ? "xaxis.labelrotation"
                        : "yaxis.labelrotation"] = value;
                    break;

                default:
                    // Forward unknown keys verbatim; SetChartProperties will flag them as unsupported.
                    translated[key] = value;
                    break;
            }
        }

        return SetChartProperties(chartPart, translated);
    }
}
