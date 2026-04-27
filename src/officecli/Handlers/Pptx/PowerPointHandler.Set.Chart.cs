// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

// Per-element-type Set helpers for chart paths. Mechanically extracted
// from the original god-method Set(); each helper owns one path-pattern's
// full handling. No behavior change.
public partial class PowerPointHandler
{
    private List<string> SetChartAxisByPath(Match chartAxisSetMatch, Dictionary<string, string> properties)
    {
        var caSlideIdx = int.Parse(chartAxisSetMatch.Groups[1].Value);
        var caChartIdx = int.Parse(chartAxisSetMatch.Groups[2].Value);
        var caRole = chartAxisSetMatch.Groups[3].Value;
        var (caSlidePart, _, caChartPart, _) = ResolveChart(caSlideIdx, caChartIdx);
        if (caChartPart == null)
            throw new ArgumentException($"Axis Set not supported on extended charts.");
        var axUnsupported = ChartHelper.SetAxisProperties(caChartPart, caRole, properties);
        GetSlide(caSlidePart).Save();
        return axUnsupported;
    }

    private List<string> SetChartByPath(Match chartSetMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(chartSetMatch.Groups[1].Value);
        var chartIdx = int.Parse(chartSetMatch.Groups[2].Value);
        var seriesIdx = chartSetMatch.Groups[3].Success ? int.Parse(chartSetMatch.Groups[3].Value) : 0;

        var (slidePart, chartGf, chartPart, extChartPart) = ResolveChart(slideIdx, chartIdx);

        // If series sub-path, prefix all properties with series{N}. for ChartSetter
        var chartProps = new Dictionary<string, string>();
        var gfProps = new Dictionary<string, string>();
        if (seriesIdx > 0)
        {
            foreach (var (key, value) in properties)
                chartProps[$"series{seriesIdx}.{key}"] = value;
        }
        else
        {
            foreach (var (key, value) in properties)
            {
                if (key.ToLowerInvariant() is "x" or "y" or "width" or "height" or "name")
                    gfProps[key] = value;
                else
                    chartProps[key] = value;
            }
        }

        // Position/size
        foreach (var (key, value) in gfProps)
        {
            switch (key.ToLowerInvariant())
            {
                case "x" or "y" or "width" or "height":
                {
                    var xfrm = chartGf.Transform ?? (chartGf.Transform = new Transform());
                    TryApplyPositionSize(key.ToLowerInvariant(), value,
                        xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                        xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                    break;
                }
                case "name":
                    var nvPr = chartGf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                    if (nvPr != null) nvPr.Name = value;
                    break;
            }
        }

        List<string> unsupported;
        if (chartPart != null)
        {
            unsupported = ChartHelper.SetChartProperties(chartPart, chartProps);
        }
        else if (extChartPart != null)
        {
            // cx:chart — delegates to ChartExBuilder.SetChartProperties.
            // Same shared implementation as Excel/Word.
            unsupported = ChartExBuilder.SetChartProperties(extChartPart, chartProps);
        }
        else
        {
            unsupported = chartProps.Keys.ToList();
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }
}
