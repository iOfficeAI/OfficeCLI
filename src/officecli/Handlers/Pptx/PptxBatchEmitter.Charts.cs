// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(chart-data-string): mirrors WordBatchEmitter.Charts.cs —
    // emit a semantic `data="Name1:v1,v2;Name2:v3,v4"` string reconstructed
    // from series children that AddChart re-builds at replay. The embedded
    // xlsx (ppt/embeddings/Microsoft_Excel_Worksheet.xlsx) is lossy on
    // round-trip: formulas, conditional formatting, defined names from the
    // source workbook are dropped. Same trade-off as docx — chart visual
    // round-trips, chart workbook does not.
    private static void EmitChart(PowerPointHandler ppt, DocumentNode chartNode,
                                  string parentSlidePath, List<BatchItem> items,
                                  SlideEmitContext ctx, int chartOrdinal)
    {
        // depth=1 so series children materialize with their name/values.
        var fullChart = ppt.Get(chartNode.Path, depth: 1);
        var props = FilterEmittableProps(fullChart.Format);
        // Strip Get-only keys AddChart neither expects nor accepts.
        props.Remove("id");
        props.Remove("seriesCount");

        // Scatter/bubble charts intrinsically carry TWO c:valAx (X and Y are
        // both value-axes — no category axis), which the Reader's
        // "multi-valAx ⇒ secondary axis" heuristic mistakes for a combo
        // primary+secondary pair. Re-emitting `secondaryAxis=1,2` on a scatter
        // forces ApplySecondaryAxis at replay, which retags one series's axIds
        // and (because plotArea now contains two ScatterCharts with disjoint
        // series binds) gets detected as `chartType=combo` on the next Get.
        // Drop the spurious key for these chart types — primary/secondary is
        // not a meaningful concept on a scatter or bubble plot.
        if (props.TryGetValue("chartType", out var chartTypeStr)
            && (chartTypeStr.Equals("scatter", StringComparison.OrdinalIgnoreCase)
                || chartTypeStr.Equals("bubble", StringComparison.OrdinalIgnoreCase)))
        {
            props.Remove("secondaryAxis");
        }

        // Reconstruct AddChart's data="Name1:v1,v2;..." input from the
        // series children (each carries `name` + `values` Format keys).
        var seriesParts = new List<string>();
        if (fullChart.Children != null)
        {
            foreach (var s in fullChart.Children)
            {
                if (s.Type != "series") continue;
                // Reference-line overlay series are rebuilt at replay via the
                // chart-level `referenceLine=value:color:label:dash` prop,
                // not as a data series. Skip them from data= so the area /
                // bar / column primary chart isn't padded with a flat 60,60
                // ghost series at replay.
                if (s.Format.TryGetValue("refLine", out var rl)
                    && (rl?.ToString() ?? "").Equals("true", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!s.Format.TryGetValue("name", out var nObj) || nObj == null) continue;
                if (!s.Format.TryGetValue("values", out var vObj) || vObj == null) continue;
                var name = nObj.ToString() ?? "";
                var vals = vObj.ToString() ?? "";
                if (name.Length == 0 || vals.Length == 0) continue;
                seriesParts.Add($"{name}:{vals}");
            }
        }

        // Waterfall round-trip: BuildWaterfallChart encodes the user's delta
        // input into 3 stacked-bar series (Base/Increase/Decrease) with
        // cumulative values. Re-feeding those 3 series on replay doubles the
        // running total — Builder would re-encode the already-encoded data.
        // Reverse the encoding here: each category's delta is `inc[i]` if
        // inc[i] != 0 else `-dec[i]`; emit a single series under the chart's
        // own name (or "Series 1") so AddChart's waterfall path takes over.
        // Per-category names are recovered from `categories=`.
        if (props.TryGetValue("chartType", out var ctype)
            && ctype.Equals("waterfall", StringComparison.OrdinalIgnoreCase)
            && fullChart.Children != null)
        {
            var byName = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in fullChart.Children)
            {
                if (s.Type != "series") continue;
                if (!s.Format.TryGetValue("name", out var nObj)) continue;
                if (!s.Format.TryGetValue("values", out var vObj)) continue;
                var nm = nObj?.ToString() ?? "";
                var vs = (vObj?.ToString() ?? "")
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => double.TryParse(t.Trim(),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out var d) ? d : 0.0)
                    .ToArray();
                byName[nm] = vs;
            }
            if (byName.TryGetValue("Increase", out var inc)
                && byName.TryGetValue("Decrease", out var dec)
                && inc.Length == dec.Length
                && inc.Length > 0)
            {
                var deltas = new double[inc.Length];
                for (int i = 0; i < inc.Length; i++)
                    deltas[i] = inc[i] != 0 ? inc[i] : -dec[i];
                var deltaStr = string.Join(",",
                    deltas.Select(d => d.ToString("G",
                        System.Globalization.CultureInfo.InvariantCulture)));
                seriesParts = new List<string> { $"Waterfall:{deltaStr}" };
                // Strip per-series color/style props that referred to the
                // encoded triplet — Builder re-applies increase/decrease/
                // total colors from explicit chart-level keys.
            }
        }

        if (seriesParts.Count > 0)
            props["data"] = string.Join(";", seriesParts);

        // Per-series style round-trip: NodeBuilder emits color/lineWidth/
        // lineDash/marker/smooth on each series child Format, but the chart-
        // level `add` step has no series sub-nodes to attach those to. The
        // chart Setter accepts dotted per-series keys (series{N}.color,
        // series{N}.lineWidth, ...) — re-flatten them here so a chart with
        // an explicit series color (#C00000 darkred etc.) round-trips
        // instead of falling back to the DefaultSeriesColors palette.
        // Skipped for waterfall (handled via increase/decrease/totalColor
        // chart-level props; emitting series1.color=transparent for "Base"
        // would fight Builder's NoFill encoding).
        var isWaterfall = props.TryGetValue("chartType", out var ctForSeries)
            && ctForSeries.Equals("waterfall", StringComparison.OrdinalIgnoreCase);
        if (!isWaterfall && fullChart.Children != null)
        {
            int seriesIdx = 0;
            bool anySeriesTrendline = false;
            foreach (var s in fullChart.Children)
            {
                if (s.Type != "series") continue;
                if (s.Format.TryGetValue("refLine", out var rlFlag)
                    && (rlFlag?.ToString() ?? "").Equals("true", StringComparison.OrdinalIgnoreCase))
                    continue; // ref-line overlay rebuilt via chart-level referenceLine=
                seriesIdx++;
                foreach (var key in new[] { "color", "lineWidth", "lineDash",
                    "marker", "markerSize", "smooth", "outlineColor",
                    "outlineWidth", "outlineDash", "transparency", "gradient",
                    "trendline", "trendline.dispRSqr", "trendline.dispEq",
                    "errbars" })
                {
                    if (s.Format.TryGetValue(key, out var val) && val != null)
                    {
                        var sval = val.ToString();
                        if (string.IsNullOrEmpty(sval)) continue;
                        // Idempotence: when the chart-level `key` (Reader's
                        // first-series summary, replayed at chart Setter time
                        // by fanning to every series) already carries the
                        // exact same value, the per-series row would be a
                        // no-op on replay but would surface as a phantom
                        // `series{N}.{key}=…` row in dump output that wasn't
                        // present in the source OOXML. Skip the duplicate.
                        // Trendline is excluded — its chart-level form is
                        // stripped below when any series trendline is emitted,
                        // and the per-series row carries spec details (type/
                        // order) that the chart-level summary loses.
                        if (key != "trendline"
                            && props.TryGetValue(key, out var chartLevel)
                            && string.Equals(chartLevel, sval, StringComparison.Ordinal))
                            continue;
                        props[$"series{seriesIdx}.{key}"] = sval;
                        if (key == "trendline") anySeriesTrendline = true;
                    }
                }
            }
            // Chart-level `trendline` is Reader's first-series summary — once
            // per-series `seriesN.trendline` rows have been emitted, the
            // chart-level key would replay through BuildTrendline a SECOND
            // time on series 1, doubling its trendline collection (or, for
            // multi-series varied trendlines, would overwrite series 1's
            // exp/log spec with the first-scanned type). Strip it.
            if (anySeriesTrendline)
            {
                props.Remove("trendline");
                props.Remove("trendline.dispRSqr");
                props.Remove("trendline.dispEq");
            }
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "chart",
            Props = props.Count > 0 ? props : null,
        });

        // Axis-role round-trip. EmitChart's add row covers chart-level axis
        // shortcuts (axismin/axismax/axistitle) that only target the primary
        // value axis — for any per-role override (especially role=value2 on a
        // secondary axis, or category/series tick/format tweaks) the add path
        // is silent. Re-query each axis via the handler's Get on the axis
        // sub-path (same path the Set side accepts) and emit a
        // `set /slide[N]/chart[K]/axis[@role=ROLE]` row carrying the non-default
        // keys. Missing roles (pie / doughnut / treemap have no axes) are
        // silently skipped.
        // Pass BOTH paths: the @id= path (or whatever Get returned) for
        // reading axes on the live file, and the positional `chart[K]` path
        // for the emitted `set` row — at replay the chart is freshly added
        // and gets a NEW cNvPr.Id, so an @id= selector from the source file
        // would no longer match. Positional index is stable (chart-only,
        // 1-based within the slide, same ordering as ResolveChart).
        var replayChartPath = $"{parentSlidePath}/chart[{chartOrdinal}]";
        EmitChartAxesIfDifferent(ppt, fullChart.Path, replayChartPath, items);
    }

    // Keys BuildAxisNode emits with synthetic defaults that match a freshly-
    // built axis under AddChart — emitting them would be a no-op and just add
    // noise to dump output. `role` is the selector key, not a property.
    private static readonly HashSet<string> AxisDefaultSkipKeys =
        new(StringComparer.OrdinalIgnoreCase) { "role" };

    // Value tuples (key, defaultValue) — skip the key on emit when the value
    // matches the AddChart default. Avoids re-emitting `majorTickMark=out`,
    // `crosses=autoZero` etc. on every dump even when the user never touched
    // the axis. Defaults sourced from ChartHelper.Builder PostAxisDefaults
    // (see SetterHelpers — bar/column/line all share these).
    private static readonly Dictionary<string, string> AxisDefaultValueSkips =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["majorTickMark"] = "out",
            ["minorTickMark"] = "none",
            ["tickLabelPos"] = "nextTo",
            ["crosses"] = "autoZero",
            ["crossBetween"] = "between",
            ["majorGridlines"] = "true",   // value axis seeds gridlines on
            ["minorGridlines"] = "false",
        };

    // For role=value, BuildChartSpace's Reader exposes the same axis values
    // as chart-level shortcuts (axisTitle / axisMin / axisMax / majorUnit /
    // axisNumFmt). EmitChart already routes those through the chart's `add`
    // row. Re-emitting them as `set axis[@role=value] title=` / `min=` /
    // ... runs the Setter twice — second pass nukes the title's run-styles
    // (RemoveAllChildren<Title>+rebuild without preserving font/color/bold),
    // and burns extra batch rows. Strip the duplicates from the role=value
    // axis emit.
    private static readonly HashSet<string> ValueAxisChartLevelKeys =
        new(StringComparer.OrdinalIgnoreCase) { "title", "min", "max", "majorUnit", "format" };

    private static void EmitChartAxesIfDifferent(PowerPointHandler ppt,
        string readChartPath, string replayChartPath, List<BatchItem> items)
    {
        if (string.IsNullOrEmpty(readChartPath)) return;

        foreach (var role in new[] { "category", "value", "value2", "series" })
        {
            var readAxisPath = $"{readChartPath}/axis[@role={role}]";
            var replayAxisPath = $"{replayChartPath}/axis[@role={role}]";
            DocumentNode? axisNode;
            try { axisNode = ppt.Get(readAxisPath); }
            catch { continue; } // axis missing on this chart type — skip

            var setProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var isPrimaryValueAxis = role == "value";
            foreach (var (k, v) in axisNode.Format)
            {
                if (v == null) continue;
                if (AxisDefaultSkipKeys.Contains(k)) continue;
                if (isPrimaryValueAxis && ValueAxisChartLevelKeys.Contains(k)) continue;
                // Skip BuildAxisNode "always-emit" gridline booleans when
                // false — AddChart leaves the default off, so emitting
                // `majorGridlines=false` is a no-op that adds noise.
                var s = v.ToString();
                if (string.IsNullOrEmpty(s)) continue;
                if (AxisDefaultValueSkips.TryGetValue(k, out var def)
                    && string.Equals(s, def, StringComparison.OrdinalIgnoreCase))
                    continue;
                // Gridlines on/off (and color/width siblings) are owned by the
                // chart-level `add` row — emitting `set axis minorGridlines=true`
                // here causes ChartHelper.Setter to RemoveAllChildren<MinorGridlines>
                // and re-add a bare one, nuking the color/width siblings that
                // the chart-level minorGridlineColor / minorGridlineWidth props
                // had just installed. Drop both boolean keys from axis emit
                // entirely; the chart-level keys round-trip the full state.
                if (k.Equals("majorGridlines", StringComparison.OrdinalIgnoreCase)
                  || k.Equals("minorGridlines", StringComparison.OrdinalIgnoreCase))
                    continue;
                // visible=true is BuildAxisNode's "always-emit" default.
                if (k.Equals("visible", StringComparison.OrdinalIgnoreCase)
                    && s!.Equals("true", StringComparison.OrdinalIgnoreCase))
                    continue;
                setProps[k] = s!;
            }
            if (setProps.Count == 0) continue;

            items.Add(new BatchItem
            {
                Command = "set",
                Path = replayAxisPath,
                Props = setProps,
            });
        }
    }
}
