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

        // CONSISTENCY(anchor-shorthand): schemas/help/_shared/chart.pptx-xlsx.json
        // declares anchor as add+set with example `anchor=2cm,3cm,18cm,10cm`
        // for pptx (vs `anchor=D2:J18` cell-range form for xlsx). Expand the
        // 4-tuple shorthand into x/y/w/h so the existing position handling
        // below picks them up. Series-sub-path Set has no position concept,
        // so anchor is silently ignored there (same as x/y/w/h would be).
        if (seriesIdx == 0
            && properties.TryGetValue("anchor", out var anchorRaw)
            && !string.IsNullOrWhiteSpace(anchorRaw))
        {
            var parts = anchorRaw.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 4)
                throw new ArgumentException(
                    $"Invalid pptx chart anchor '{anchorRaw}'. Expected 'x,y,w,h' (e.g. '2cm,3cm,18cm,10cm'). For xlsx use a cell range like 'D2:J18'.");
            // Override any explicitly-supplied x/y/w/h so single-prop intent is
            // unambiguous: anchor wins because the user picked the compound form.
            gfProps["x"] = parts[0];
            gfProps["y"] = parts[1];
            gfProps["width"] = parts[2];
            gfProps["height"] = parts[3];
        }

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
                {
                    if (!gfProps.ContainsKey(key)) gfProps[key] = value;
                }
                else if (key.Equals("anchor", StringComparison.OrdinalIgnoreCase))
                    continue; // already expanded into gfProps above
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
