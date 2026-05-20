# Scatter Charts Showcase

This demo consists of three files that work together:

- **charts-scatter.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-scatter.xlsx** — The generated workbook with 7 sheets (1 default + 6 chart sheets, 24 charts total).
- **charts-scatter.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-scatter.py
# → charts-scatter.xlsx
```

## Chart Sheets

### Sheet: 1-Scatter Fundamentals

Four scatter variants covering markers+lines, marker-only, smooth curves, and line-only.

```bash
# Basic scatter with circle markers and connecting lines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop series1="Male:62,68,72,78,82,88,95" \
  --prop categories=160,165,170,175,180,185,190 \
  --prop marker=circle --prop markerSize=6 --prop lineWidth=1.5 \
  --prop catTitle=Height (cm) --prop axisTitle=Weight (kg)

# Scatter marker-only (no connecting lines)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter --prop scatterStyle=marker \
  --prop markerSize=8 --prop gridlines=D9D9D9:0.5:dot

# Scatter smooth curve (Bezier interpolation)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter --prop scatterStyle=smooth \
  --prop smooth=true --prop marker=diamond --prop lineWidth=2

# Scatter line-only (no markers)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter --prop scatterStyle=line \
  --prop showMarker=false --prop lineWidth=2.5 --prop lineDash=dash
```

**Features:** `scatter`, `scatterStyle=marker|smooth|line`, `smooth=true`, `marker=circle|diamond`, `markerSize`, `lineWidth`, `lineDash=dash`, `showMarker=false`, `catTitle`, `axisTitle`, `gridlines`

### Sheet: 2-Marker Styles

Four charts demonstrating all marker shapes and per-series marker control.

```bash
# Per-series markers: circle, diamond, square
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop series1.marker=circle --prop series2.marker=diamond \
  --prop series3.marker=square --prop markerSize=8

# Per-series markers: triangle, star, x
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop series1.marker=triangle --prop series2.marker=star \
  --prop series3.marker=x --prop markerSize=9

# Large markers with plus and dash shapes
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter --prop scatterStyle=marker \
  --prop series1.marker=circle --prop series2.marker=plus \
  --prop series3.marker=dash --prop markerSize=10

# showMarker=false with lineDash=dashDot
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter --prop scatterStyle=lineMarker \
  --prop showMarker=false --prop lineDash=dashDot
```

**Features:** `series{N}.marker=circle|diamond|square|triangle|star|x|plus|dash`, `markerSize`, `scatterStyle=lineMarker|marker`, `showMarker=false`, `lineDash=dashDot`

### Sheet: 3-Trendlines

Four charts covering all trendline types and sub-properties.

```bash
# Linear trendline with equation display
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter --prop scatterStyle=marker \
  --prop trendline=linear \
  --prop series1.trendline.equation=true

# Polynomial (order 3) with R-squared display
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter --prop scatterStyle=marker \
  --prop trendline=poly:3 \
  --prop series1.trendline.rsquared=true

# Exponential with forward/backward extrapolation
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter --prop scatterStyle=marker \
  --prop trendline=exp:2:1 \
  --prop series1.trendline.name=Exponential Fit

# Per-series trendlines: linear vs logarithmic
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter --prop scatterStyle=marker \
  --prop series1.trendline=linear --prop series2.trendline=log \
  --prop series1.trendline.equation=true \
  --prop series2.trendline.rsquared=true
```

**Features:** `trendline=linear|poly:N|exp|log|power|movingAvg`, `trendline=exp:forward:backward` (extrapolation), `series{N}.trendline` (per-series), `series{N}.trendline.equation`, `series{N}.trendline.rsquared`, `series{N}.trendline.name`

### Sheet: 4-Error Bars

Four charts covering all error bar types on scatter series.

```bash
# Fixed error bars (+/-5)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop errBars=fixed:5

# Percentage error bars (10%)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop errBars=percent:10

# Standard deviation error bars
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop errBars=stddev

# Standard error with series shadow
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop errBars=stderr \
  --prop series.shadow=000000-4-315-2-30
```

**Features:** `errBars=fixed:N|percent:N|stddev|stderr`, `series.shadow`

### Sheet: 5-Styling

Four charts covering title styling, fills, gradients, borders, and axis formatting.

```bash
# Title styling with series shadow and outline
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop title.shadow=000000-3-315-2-30 \
  --prop series.shadow=000000-4-315-2-30 \
  --prop series.outline=333333-1.5

# Gradients, transparency, plotFill, chartFill
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90' \
  --prop transparency=20 \
  --prop plotFill=F5F5F5 --prop chartFill=FAFAFA

# Axis font, gridlines, minor gridlines, axis line
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop axisfont=9:C00000:Arial \
  --prop gridlines=BFBFBF:0.75:solid \
  --prop minorGridlines=E0E0E0:0.25:dot \
  --prop axisLine=333333:1

# Chart/plot borders and rounded corners
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop chartArea.border=333333-1.5 \
  --prop plotArea.border=999999-0.75 \
  --prop roundedCorners=true
```

**Features:** `title.font/size/color/bold`, `title.shadow`, `series.shadow`, `series.outline`, `gradients`, `transparency`, `plotFill`, `chartFill`, `axisfont`, `gridlines`, `minorGridlines`, `axisLine`, `chartArea.border`, `plotArea.border`, `roundedCorners`

### Sheet: 6-Advanced

Four charts covering secondary axis, reference lines, log scale, and conditional coloring.

```bash
# Secondary Y-axis for dual-unit scatter
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop secondaryAxis=2

# Reference line (horizontal target)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop referenceLine=75:FF0000:Target:dash

# Logarithmic axis with min/max bounds
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop logBase=10 \
  --prop axisMin=1 --prop axisMax=10000

# Data labels with conditional color rule
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter --prop scatterStyle=marker \
  --prop dataLabels=true --prop labelPos=top \
  --prop colorRule=60:C00000:00AA00
```

**Features:** `secondaryAxis`, `referenceLine=value:color:label:dash`, `logBase`, `axisMin`, `axisMax`, `dataLabels`, `labelPos=top`, `colorRule=threshold:belowColor:aboveColor`

## Inspect the Generated File

```bash
officecli query charts-scatter.xlsx chart
officecli get charts-scatter.xlsx "/1-Scatter Fundamentals/chart[1]"
```
