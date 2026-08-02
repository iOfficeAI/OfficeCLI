// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    internal static Dictionary<string, string> BuildChartProps(ChartSpec spec)
    {
        // AddChart ingests data series via a single `data="Name1:v1,v2;Name2:v1,v2"`
        // string. Reconstruct that string from the series children Get
        // exposes; categories come from the chart's own Format key.
        var props = FilterEmittableProps(spec.Format);
        // CONSISTENCY(internal-roundtrip-keys): pull chart-level verbatim
        // XML metadata (axisTitle.pPr / catTitle.pPr) off InternalFormat
        // so the dump→replay path still carries them. They live off the
        // public Format dict to keep user-facing Get output free of raw
        // OOXML strings, but the Setter consumes them post-build to
        // restore axis-title paragraph styling.
        foreach (var (ik, iv) in spec.InternalFormat)
        {
            if (iv == null) continue;
            var ivs = iv.ToString();
            if (string.IsNullOrEmpty(ivs)) continue;
            if (!props.ContainsKey(ik)) props[ik] = ivs;
        }
        // Strip Get-only / SDK-managed keys that AddChart neither expects
        // nor accepts.
        props.Remove("id");
        props.Remove("seriesCount");

        // BUG-DUMP-R34-1: the chart-level verbatim axis-line / plot-area spPr
        // fragments (valAx.spPr / catAx.spPr / plotArea.spPr) are the source of
        // truth for the value-axis line, category-axis line, and plot-area
        // border/fill. When present, drop the lossy granular keys they
        // supersede so replay doesn't re-derive a second, conflicting outline
        // or fill on top of the verbatim XML:
        //   - plotArea.spPr carries the gradFill AND the border <a:ln> ->
        //     drop plotFill + plotArea.border[.*].
        //   - valAx.spPr / catAx.spPr carry the axis line -> drop
        //     valAxisLine[.*] / catAxisLine[.*].
        // The verbatim keys themselves flow through to SetChartProperties (they
        // are deferred + handled there). Plain charts emit none of these, so
        // this is a no-op for them.
        if (props.ContainsKey("plotArea.spPr"))
        {
            props.Remove("plotFill");
            foreach (var k in props.Keys
                .Where(k => k.StartsWith("plotArea.border", StringComparison.OrdinalIgnoreCase))
                .ToList())
                props.Remove(k);
        }
        if (props.ContainsKey("valAx.spPr"))
            foreach (var k in props.Keys
                .Where(k => k.StartsWith("valAxisLine", StringComparison.OrdinalIgnoreCase))
                .ToList())
                props.Remove(k);
        if (props.ContainsKey("catAx.spPr"))
            foreach (var k in props.Keys
                .Where(k => k.StartsWith("catAxisLine", StringComparison.OrdinalIgnoreCase))
                .ToList())
                props.Remove(k);
        // gridline.spPr / minorGridline.spPr carry the gridline outline verbatim
        // (tx1+lumMod tint, cap/cmpd/join) — drop the lossy granular
        // gridlineColor/Width/Dash they supersede so replay doesn't re-derive a
        // solid black line on top of the verbatim XML.
        if (props.ContainsKey("gridline.spPr"))
            foreach (var k in new[] { "gridlineColor", "gridlineWidth", "gridlineDash" })
                props.Remove(k);
        if (props.ContainsKey("minorGridline.spPr"))
            foreach (var k in new[] { "minorGridlineColor", "minorGridlineWidth", "minorGridlineDash" })
                props.Remove(k);
        // chartArea.spPr carries the chart frame fill + border verbatim (tx1+lumMod
        // tint) -> drop the lossy chartFill + chartArea.border[.*] it supersedes.
        if (props.ContainsKey("chartArea.spPr"))
        {
            props.Remove("chartFill");
            foreach (var k in props.Keys
                .Where(k => k.StartsWith("chartArea.border", StringComparison.OrdinalIgnoreCase))
                .ToList())
                props.Remove(k);
        }

        // BUG-DUMP-R35-1: the verbatim per-axis text fragments (catAx.txPr /
        // valAx.txPr) supersede the single `axisFont` key, which read only the
        // VALUE axis and on rebuild applied that one font to BOTH axes —
        // clobbering the category axis's distinct font. Drop `axisFont` once a
        // verbatim axis txPr is present so the two fragments aren't fighting one
        // chart-level font. (title.pPr is purely ADDITIVE — it restores the
        // title's defRPr colour/typeface/alignment that the run-level title.*
        // keys never carried — so no title key is dropped.) Plain charts emit
        // none of these, so this is a no-op for them.
        if (props.ContainsKey("catAx.txPr") || props.ContainsKey("valAx.txPr"))
            props.Remove("axisFont");

        // Build data="Name:v1,v2;..." from series children, plus the per-series
        // dotted props AddChart honors (series{N}.color, series{N}.dataLabels).
        // N is 1-based over the EMITTED (non-reference-line) series, matching how
        // AddChart maps the dotted keys onto the series it parses from `data=`.
        // Without this, a bar/column chart whose series carry explicit fill
        // colors (Q1 blue / Q2 orange / Q3 green) round-tripped to Word's default
        // single-color palette (all blue), and per-series data labels (showVal)
        // were dropped. Reader already surfaces series Format["color"] and
        // Format["dataLabels"]; only this emit was missing.
        var seriesParts = new List<string>();
        int seriesIdx = 0;
        bool emittedPerSeriesFill = false;
        bool emittedPerSeriesScalar = false;
        var dataLabelFlags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var s in spec.Series)
        {
            if (s.Type != "series") continue;
            // Skip reference-line series: AddReferenceLine re-creates the Target
            // series from `referenceLine=...` props. Including its values in the
            // data string would duplicate the series on replay.
            if (s.Format.TryGetValue("refLine", out var rl) && rl?.ToString() == "true") continue;
            if (!s.Format.TryGetValue("name", out var nObj) || nObj == null) continue;
            if (!s.Format.TryGetValue("values", out var vObj) || vObj == null) continue;
            var name = nObj.ToString() ?? "";
            var vals = vObj.ToString() ?? "";
            if (name.Length == 0 || vals.Length == 0) continue;
            seriesParts.Add($"{name}:{vals}");
            seriesIdx++;

            // BUG-DUMP-R33-1: when the Reader captured the series' styling
            // sub-elements verbatim, forward them as `series{N}.spPr/dPt/dLbls`
            // and SUPPRESS the granular keys they subsume. The verbatim XML is
            // authoritative — re-applying the lossy granular color/gradient/
            // outlineColor/shadow on top would either double-fill or fight the
            // captured spPr; the per-point color key is fully replaced by the
            // verbatim dPt list. Without this gate the round-trip kept dropping
            // <a:ln>/<a:outerShdw>/<c:dPt>/rich <c:dLbls>.
            // CONSISTENCY(internal-roundtrip-keys): series spPr verbatim XML
            // lives on InternalFormat (kept off public Format so user-facing
            // Get output doesn't expose raw OOXML). Reader writes there;
            // emitters read from there.
            bool hasVerbatimSpPr = s.InternalFormat.TryGetValue("spPr", out var spObj) && spObj is string sp && sp.Length > 0;
            // R45: dPt/dLbls verbatim XML moved to InternalFormat (was Format).
            // Reader writes there; emitter must read there too.
            bool hasVerbatimDpt = s.InternalFormat.TryGetValue("dPt", out var dpObj) && dpObj is string dp && dp.Length > 0;
            bool hasVerbatimDLbls = s.InternalFormat.TryGetValue("dLbls", out var dlbObj) && dlbObj is string dlb && dlb.Length > 0;
            if (hasVerbatimSpPr)
            {
                props[$"series{seriesIdx}.spPr"] = (string)spObj!;
                emittedPerSeriesFill = true;
            }
            if (hasVerbatimDpt)
                props[$"series{seriesIdx}.dPt"] = (string)dpObj!;
            if (hasVerbatimDLbls)
                props[$"series{seriesIdx}.dLbls"] = (string)dlbObj!;

            // Per-series explicit fill color (srgbClr / "none"). Reader only sets
            // this key when the series carries an explicit <c:spPr> fill, so its
            // presence is meaningful — forward it verbatim. Skip when a verbatim
            // spPr already carries the fill (it is the source of truth).
            if (!hasVerbatimSpPr && s.Format.TryGetValue("color", out var cObj) && cObj != null)
            {
                var c = cObj.ToString() ?? "";
                if (c.Length > 0) { props[$"series{seriesIdx}.color"] = c; emittedPerSeriesFill = true; }
            }
            // Per-series gradient fill (e.g. "5B9BD5-2E75B6:90:s0"). Distinct
            // from a chart-level `gradient=` which applies one gradient to ALL
            // series — here each series can carry its own (Q1 blue / Q2 orange).
            else if (!hasVerbatimSpPr && s.Format.TryGetValue("gradient", out var gObj) && gObj != null)
            {
                var g = gObj.ToString() ?? "";
                if (g.Length > 0) { props[$"series{seriesIdx}.gradient"] = g; emittedPerSeriesFill = true; }
            }
            // Per-series scalar styling (marker/smooth/trendline/errBars…).
            // The chart-level keys for these are series-1 REPRESENTATIVES; a
            // chart whose series diverge (series1 circle, series2 square) was
            // collapsed to series1's value on every series at replay. Emit the
            // per-series dotted keys HandleSeriesDottedProperty consumes and
            // drop the chart-level representative below.
            foreach (var (fmtKey, dotKey) in SeriesScalarKeys)
            {
                if (s.Format.TryGetValue(fmtKey, out var psObj) && psObj != null
                    && psObj.ToString() is { Length: > 0 } psVal)
                {
                    props[$"series{seriesIdx}.{dotKey}"] = psVal;
                    emittedPerSeriesScalar = true;
                }
            }
            // Trendline display/forecast sub-options → chart-level fan-out keys.
            foreach (var (fmtKey, chartKey) in TrendlineSubKeys)
            {
                if (s.Format.TryGetValue(fmtKey, out var tvObj) && tvObj != null
                    && tvObj.ToString() is { Length: > 0 } tv)
                {
                    // Chart-level fan-out keys, not SeriesScalarKeys reps, so no
                    // emittedPerSeriesScalar flip needed (the base `trendline`
                    // key already sets it for any series carrying a trendline).
                    props[chartKey] = tv;
                }
            }

            // Per-series data-label show flags (value/category/series/percent).
            // AddChart has no per-series data-label setter, so accumulate the
            // union across series and emit the chart-level datalabels.show* keys
            // below. (Common case: a single-series chart showing values, or all
            // series sharing the same labels.)
            // BUG-DUMP-R33-1: skip when a verbatim dLbls was captured — its XML
            // already carries the show flags, so funnelling them into the
            // chart-level union would write a second, conflicting <c:dLbls>.
            if (!hasVerbatimDLbls && s.Format.TryGetValue("dataLabels", out var dObj) && dObj != null)
            {
                foreach (var flag in (dObj.ToString() ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                    dataLabelFlags.Add(flag);
            }
        }
        if (seriesParts.Count > 0)
        {
            props["data"] = string.Join(";", seriesParts);
        }

        // Waterfall idempotency: a waterfall chart is stored as a stacked bar
        // with three DERIVED series (Base / Increase / Decrease). DetectChartType
        // reports it as "waterfall", but the raw series children still carry the
        // derived stacks. If we replay those as `data=`, BuildWaterfallChart
        // derives base/increase AGAIN on top of the already-derived values —
        // corrupting heights and dropping the first bar. Reconstruct the single
        // logical change series (value[i] = Increase[i] - Decrease[i]) and feed
        // that as `data=` so the round-trip is idempotent. The derived triplet's
        // per-series styling keys are re-generated by the builder from
        // increaseColor / decreaseColor / totalColor (already surfaced at chart
        // level), so strip them here.
        if (props.TryGetValue("chartType", out var ctW)
            && string.Equals(ctW, "waterfall", StringComparison.OrdinalIgnoreCase)
            && seriesParts.Count == 3)
        {
            var logical = CollapseWaterfallSeries(seriesParts);
            if (logical != null)
            {
                props["data"] = logical;
                // Strip the derived-triplet series keys (series1, series1.spPr,
                // series2.dPt, series2.color, series3.color, …). The waterfall
                // builder regenerates all styling from the chart-level colors.
                foreach (var k in props.Keys
                    .Where(k => System.Text.RegularExpressions.Regex.IsMatch(
                        k, @"^series\d+(\.|$)", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    .ToList())
                    props.Remove(k);
            }
        }
        // The chart-level `gradient` (and `color`) key the Reader surfaces is a
        // chart-scope REPRESENTATIVE — it reports the FIRST series' fill. When
        // series carry distinct per-series fills, replaying that chart-level key
        // would paint EVERY series with series-1's fill (the "all blue" bug).
        // Drop it once per-series fills are emitted.
        if (emittedPerSeriesFill)
        {
            props.Remove("gradient");
            props.Remove("color");
        }
        // Same representative-key rule for the per-series scalars: once emitted
        // as series{N}.* the chart-level copies (which report series 1 only)
        // must go, or replay re-applies series-1's value to every series.
        if (emittedPerSeriesScalar)
        {
            foreach (var (fmtKey, _) in SeriesScalarKeys)
                props.Remove(fmtKey);
        }
        // Map the collected series data-label flags onto the chart-level keys
        // AddChart applies (no per-series data-label setter exists). The
        // chart-level dLbls builder now inserts the show flags in schema order.
        if (dataLabelFlags.Count > 0)
        {
            if (dataLabelFlags.Contains("value")) props["datalabels.showvalue"] = "true";
            if (dataLabelFlags.Contains("category")) props["datalabels.showcatname"] = "true";
            if (dataLabelFlags.Contains("series")) props["datalabels.showsername"] = "true";
            if (dataLabelFlags.Contains("percent")) props["datalabels.showpercent"] = "true";
        }
        return props;
    }

    /// <summary>
    /// Collapse a waterfall chart's stored Base/Increase/Decrease stacked-bar
    /// triplet back into the single logical change series the waterfall builder
    /// expects (value[i] = Increase[i] - Decrease[i]). The last bar is the
    /// auto-computed total, so its logical value is irrelevant on replay
    /// (BuildWaterfallChart recomputes it); we emit the total magnitude for a
    /// readable data string. Returns null when the triplet doesn't match the
    /// expected Base/Increase/Decrease shape.
    /// </summary>
    private static string? CollapseWaterfallSeries(List<string> seriesParts)
    {
        double[]? Parse(string part)
        {
            var colon = part.IndexOf(':');
            if (colon < 0) return null;
            var name = part[..colon];
            var vals = part[(colon + 1)..].Split(',', StringSplitOptions.RemoveEmptyEntries);
            var arr = new double[vals.Length];
            for (int i = 0; i < vals.Length; i++)
                if (!double.TryParse(vals[i], System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out arr[i]))
                    return null;
            return arr;
        }
        // Order is Base, Increase, Decrease (Builder convention).
        var inc = Parse(seriesParts[1]);
        var dec = Parse(seriesParts[2]);
        if (inc == null || dec == null || inc.Length != dec.Length || inc.Length == 0)
            return null;
        var n = inc.Length;
        var logical = new double[n];
        for (int i = 0; i < n; i++)
            logical[i] = inc[i] - dec[i];
        return "Series1:" + string.Join(",", logical.Select(v =>
            v.ToString("G", System.Globalization.CultureInfo.InvariantCulture)));
    }

    // Reader series Format key -> AddChart series{N}.<dotted> key for the
    // scalar styling props that would otherwise collapse to series 1's value.
    private static readonly (string, string)[] SeriesScalarKeys =
    {
        ("marker", "marker"),
        ("markerSize", "markersize"),
        ("markerColor", "markercolor"),
        ("smooth", "smooth"),
        ("trendline", "trendline"),
        ("errBars", "errbars"),
    };

    // Trendline display/forecast sub-options the Reader surfaces per-series but
    // AddChart consumes only as chart-level fan-out keys. SeriesScalarKeys
    // carries just the type:order spec, so these were dropped on dump→replay.
    // Map each Reader key to the AddChart chart-level key. One trendline is the
    // common case; if multiple series carry divergent flags the last one seen
    // wins (matches the Reader's first-trendline display-flag model).
    private static readonly (string, string)[] TrendlineSubKeys =
    {
        ("trendline.dispRSqr", "trendline.displayrsquared"),
        ("trendline.dispEq", "trendline.displayequation"),
        ("trendline.forward", "trendline.forecastforward"),
        ("trendline.backward", "trendline.forecastbackward"),
        ("trendline.intercept", "trendline.intercept"),
        ("trendline.name", "trendline.label"),
    };
}
