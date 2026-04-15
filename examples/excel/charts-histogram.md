# Histogram Charts — Grand Showcase

The most thorough histogram demo officecli can produce. Every binning knob,
every styling vocabulary, every canonical distribution shape, six design
themes, four font-family type specimens, and a cohesive production-grade
ML dashboard.

This demo is three files that work together:

- **charts-histogram.py** — Python script that calls `officecli` to generate
  the workbook. Each chart command is shown as a copyable shell command in
  the comments.
- **charts-histogram.xlsx** — The generated workbook: 6 sheets, 29 charts.
- **charts-histogram.md** — This file. Maps each sheet to the features it
  demonstrates and lists the full histogram property vocabulary.

## Regenerate

```bash
cd examples/excel
python3 charts-histogram.py
# → charts-histogram.xlsx
```

## Why a dedicated histogram showcase?

Histograms are Excel's cx-namespace "extended" chart type. The binning layer
(`layoutPr/binning`) is where all the interesting knobs live — auto vs
explicit count, bin width, interval-closed side, outlier cut-offs — and
getting them right takes some care because Excel rejects the file entirely
if the XML uses the wrong form of `cx:binCount` / `cx:binSize`.

Beyond binning, the cx pipeline in officecli has full parity with regular
cChart for typography, axis scaling, area fills/borders, drop shadows,
data labels, and legend styling. This file exercises every binning knob
AND every styling knob in one place, so you can copy-paste from whichever
row most matches the shape you want.

## Sheets at a glance

| Sheet | Charts | What it demonstrates |
|---|---|---|
| 0-Hero | 1 | Full-bleed magazine-grade poster using EVERY knob |
| 1-Binning Lab | 6 | Every binning strategy on one dataset, identical styling |
| 2-Distribution Zoo | 6 | Six canonical real-world distribution shapes |
| 3-Theme Gallery | 6 | Six complete design themes on the SAME dataset |
| 4-Typography | 4 | Four font-family type specimens |
| 5-ML Dashboard | 6 | Cohesive "Production ML Model Report" dashboard |

## Sheet 0: 0-Hero

One full-bleed 27×38-cell hero chart that combines EVERY histogram knob
into a single presentation-grade poster. Dark "Midnight Academia" palette
— navy plot area, gold bars, cream title, soft grid lines, locked Y axis,
dropped shadows on both title and series, data labels with number format,
top legend with compound font styling. If this chart renders correctly,
the entire histogram pipeline is healthy.

```bash
officecli add charts-histogram.xlsx "/0-Hero" --type chart \
  --prop chartType=histogram \
  --prop title="The Shape of Data · 200-sample bell curve" \
  --prop title.color=F5F1E0 --prop title.size=22 --prop title.bold=true \
  --prop title.font="Helvetica Neue" \
  --prop "title.shadow=000000-8-45-4-70" \
  --prop series1="Samples:<200 bell values>" \
  --prop binCount=24 --prop intervalClosed=l \
  --prop fill=F0C96A --prop "series.shadow=000000-8-45-4-60" \
  --prop axismin=0 --prop axismax=28 --prop majorunit=4 \
  --prop xAxisTitle="Score" --prop yAxisTitle="Frequency" \
  --prop axisTitle.color=C9B87A --prop axisTitle.size=13 \
  --prop axisTitle.bold=true --prop axisTitle.font="Helvetica Neue" \
  --prop "axisfont=10:B8B090:Helvetica Neue" \
  --prop "axisline=6A6448:1.5" \
  --prop gridlineColor=2F3544 \
  --prop plotareafill=1A1F2C --prop "plotarea.border=3A3E4E:1.25" \
  --prop chartareafill=0B0F18 --prop "chartarea.border=2A2E3E:1" \
  --prop dataLabels=true --prop "datalabels.numfmt=0" \
  --prop legend=top --prop legend.overlay=false \
  --prop "legendfont=11:D4C994:Helvetica Neue" \
  --prop x=0 --prop y=0 --prop width=27 --prop height=38
```

**Features:** title.color / title.size / title.bold / title.font / title.shadow,
fill, series.shadow, binCount, intervalClosed, axismin/axismax/majorunit,
xAxisTitle / yAxisTitle, axisTitle.color / axisTitle.size / axisTitle.bold /
axisTitle.font, axisfont compound, axisline, gridlineColor, plotareafill,
plotarea.border, chartareafill, chartarea.border, dataLabels, datalabels.numfmt,
legend, legend.overlay, legendfont.

## Sheet 1: 1-Binning Lab

Six charts, SAME dataset (200 bell-curve samples), IDENTICAL typography and
frame — the ONLY thing that varies is the binning strategy. Put side by
side, this sheet is the binning Rosetta stone.

```bash
# 1. Auto-binning (no binCount, no binSize — Excel picks it)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram --prop series1="Samples:<values>" \
  --prop title="1 · Auto-binning (Excel default)" --prop fill=4472C4

# 2. Explicit binCount=8 (coarse)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram --prop series1="Samples:<values>" \
  --prop binCount=8 --prop title="2 · binCount=8 (coarse)"

# 3. Explicit binCount=32 (fine)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram --prop series1="Samples:<values>" \
  --prop binCount=32 --prop title="3 · binCount=32 (fine)"

# 4. Fixed bin width (binSize=5)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram --prop series1="Samples:<values>" \
  --prop binSize=5 --prop title="4 · binSize=5 (fixed-width bins)"

# 5. Outlier fencing (underflowBin=55, overflowBin=95)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram --prop series1="Samples:<values>" \
  --prop binSize=5 --prop underflowBin=55 --prop overflowBin=95

# 6. Left-closed intervals [a,b) with gapWidth=30 between bars
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram --prop series1="Samples:<values>" \
  --prop binCount=16 --prop intervalClosed=l --prop gapWidth=30
```

**Features:** `chartType=histogram`, auto-binning (default), `binCount=N`,
`binSize=W`, `underflowBin=N`, `overflowBin=M`, `intervalClosed=l`, `gapWidth=N`

Notes:
- If both `binCount` and `binSize` are given, `binCount` wins.
- Histograms default `gapWidth=0` (bars touch) to match Excel's native output.
- `intervalClosed=l` makes bins half-open `[a,b)` instead of the default `(a,b]`.
- `underflow` / `overflow` fences let the interesting bulk stay readable
  when the tail is catastrophic.

## Sheet 2: 2-Distribution Zoo

A 2×3 visual gallery of canonical real-world distribution shapes. Pattern
recognition: if you ever see one of these shapes in a telemetry chart, you
know immediately what's going on. Every chart shares the same typography
and frame; only the fill color, data, and binning strategy change.

| Shape | Data | Fill | Binning |
|---|---|---|---|
| Normal · bell curve | 200 gauss(75, 12) | #2F5597 | binCount=18 |
| Bimodal · two cohorts | 80 gauss(55,6) + 80 gauss(88,5) | #ED7D31 | binCount=22 |
| Right-skewed · log-normal | 180 exp(gauss(3.2, 0.55)) | #70AD47 | binCount=20 |
| Left-skewed · retirement | 140 75 − exp(gauss(1.6, 0.6)) | #7030A0 | binCount=18 |
| Uniform · flat floor | 160 uniform(0, 100) | #00B0F0 | binSize=10 |
| Heavy-tailed · Pareto | 200 paretovariate(1.6) × 20 | #C00000 | binSize=20, overflow=250 |

## Sheet 3: 3-Theme Gallery

Six complete design themes applied to the SAME bell-curve dataset. Each
theme is a coordinated palette: plot-area fill, chart-area fill, series
fill, gridline color, axis line color, tick-label color, title color,
title font — all chosen to read as one coherent mood.

| Theme | Mood | Plot BG | Bar | Title font |
|---|---|---|---|---|
| Midnight Academia | Dark, elegant | navy #1A1F2C | gold #F0C96A | Georgia |
| Sunset Terracotta | Warm, editorial | cream #FFF5E8 | coral #E85D4A | Georgia |
| Forest Parchment | Organic, retro | beige #F3EDD8 | forest #2F5D3A | Georgia |
| Editorial Mono | Pure grayscale | white #FFFFFF | dark #2A2A2A | Helvetica Neue |
| Neon Terminal | Cyberpunk | black #0A0A14 | cyan #00F0C8 | Courier New |
| Pastel Bloom | Soft, feminine | lavender #FDF4F8 | rose #F5A7C8 | Helvetica Neue |

Each chart uses the full parity-knob vocabulary: `plotareafill`,
`plotarea.border`, `chartareafill`, `chartarea.border`, `gridlineColor`,
`axisline`, `axisfont`, `title.color` / `title.font`, `axisTitle.color` /
`axisTitle.font`. This is the sheet to copy-paste from when you want to
build a specific look for a report.

## Sheet 4: 4-Typography

Four font-family type specimens. Same data, same geometry, nearly identical
color — only the font family varies. Side by side, this sheet shows how
typography alone can reshape a chart's tone.

| Font | Tone | Used for |
|---|---|---|
| Helvetica Neue | Modern sans | Dashboards, corporate reports |
| Georgia | Editorial serif | Magazines, long-form reports |
| Courier New | Data mono | Telemetry, engineering, terminals |
| Verdana | Friendly sans | Onboarding, public-facing UI |

Each specimen sets `title.font`, `axisTitle.font`, and the fontname segment
of the `axisfont` compound form to the same family, so the entire chart
lives in one typographic voice.

## Sheet 5: 5-ML Dashboard

A cohesive "Production ML Model Report" dashboard. Every chart wears the
same uniform — typography, frames, gridlines, axis line — but each shows
a different slice of the model's behavior, deliberately using a different
color, binning strategy, and (where relevant) outlier-fencing or axis
locking. The six read as one dashboard.

| Panel | Data shape | Color | Binning / parity knob |
|---|---|---|---|
| Inference Latency · p50–p99 | heavy-tail | #EF4444 | binSize=25, overflowBin=300, series.shadow |
| Prediction Confidence | right-skewed | #10B981 | binSize=5, axismin=0, majorunit=50 |
| Residual magnitude | half-normal | #F59E0B | binSize=0.25, intervalClosed=l |
| Token length | bimodal | #6366F1 | binCount=24 |
| GPU utilization | normal (clipped) | #8B5CF6 | binSize=5, axismin=0 axismax=50 majorunit=10 |
| Cost per request | log-normal | #EC4899 | binSize=5, overflowBin=120, dataLabels+numfmt |

This sheet shows that one typographic uniform plus per-panel color and
binning choices is enough to build a production dashboard. Copy the
`DASH` style block from `charts-histogram.py` as a starting point.

## Histogram Property Reference

| Property | Default | Notes |
|---|---|---|
| `chartType` | — | Must be `histogram` |
| `title` | — | Chart title text |
| `series1` | — | `"name:v1,v2,v3,..."` — raw values, not pre-binned |
| `binCount` | auto | Integer: force exactly N bins |
| `binSize` | auto | Number: force fixed bin width |
| `intervalClosed` | `r` | `r` = (a,b], `l` = [a,b) |
| `underflowBin` | — | Group values < N into a single `<N` bar |
| `overflowBin` | — | Group values > M into a single `>M` bar |
| `gapWidth` | `0` | Space between bars (0 = touching) |
| `fill` | — | Single-color shortcut (HEX) |
| `colors` | — | Comma list of HEX (multi-series) |
| `dataLabels` | `false` | `true` puts value count above each bar |
| `datalabels.numfmt` | — | Excel format code (`0`, `0.0`, `0.00%`, `#,##0`) |
| `xAxisTitle` / `yAxisTitle` | — | Axis titles |
| `gridlines` | `true` | Value-axis major gridlines |
| `xGridlines` | `false` | Category-axis major gridlines |
| `tickLabels` | `true` | Show bin range labels on x-axis |
| `axismin` / `axismax` | — | Value-axis range (numeric) |
| `majorunit` / `minorunit` | — | Value-axis gridline interval |
| `axis.visible` / `cataxis.visible` / `valaxis.visible` | — | Axis hidden flags |
| `axisline` | — | Axis spine: `"color"` / `"color:width"` / `"color:width:dash"` / `"none"` |
| `cataxis.line` / `valaxis.line` | — | Per-axis spine styling |
| `plotareafill` / `plotfill` | — | Plot-area solid background color |
| `plotarea.border` / `plotborder` | — | Plot-area outline |
| `chartareafill` / `chartfill` | — | Chart-area solid background color |
| `chartarea.border` / `chartborder` | — | Chart-area outline |
| `series.shadow` | — | Outer shadow on bars: `"COLOR-BLUR-ANGLE-DIST-OPACITY"` |
| `title.shadow` | — | Outer shadow on title: `"COLOR-BLUR-ANGLE-DIST-OPACITY"` |
| `legend` | — | `top` / `bottom` / `left` / `right` / `none` |
| `legend.overlay` | `false` | Legend floats on top of plot area when `true` |
| `legendfont` | — | Compound `"size:color:fontname"` |
| `title.color` / `title.size` / `title.bold` / `title.font` | — | Chart title styling |
| `axisTitle.color` / `axisTitle.size` / `axisTitle.font` / `axisTitle.bold` | — | Axis title styling (both X and Y) |
| `axisfont` | — | Compound tick-label styling: `"size:color:fontname"` |
| `gridlineColor` | — | Value-axis major gridline color |
| `xGridlineColor` | — | Category-axis major gridline color (requires `xGridlines=true`) |
| `x` / `y` / `width` / `height` | — | Chart cell placement and size |

## Inspect the Generated File

```bash
# Count all charts across all sheets
officecli query charts-histogram.xlsx chart

# Introspect a single chart's bound properties
officecli get charts-histogram.xlsx "/0-Hero/chart[1]"
officecli get charts-histogram.xlsx "/5-ML Dashboard/chart[1]"

# Render any sheet to HTML preview
officecli view charts-histogram.xlsx html > preview.html
```

> Note: officecli's HTML preview renders the full parity vocabulary
> (plot-area / chart-area fills, gridline + axis line colors, tick
> label colors, data labels, locked axis scales, gapWidth, etc.),
> but does not currently reproduce custom axis-label font families —
> all tick labels fall back to the preview's default sans font. Excel
> renders the full styling including the font family. Use the preview
> for layout + color verification, use Excel (or Numbers / LibreOffice)
> for final typographic QA.
