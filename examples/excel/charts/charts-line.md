# Line Charts Showcase

This demo consists of three files that work together:

- **charts-line.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-line.xlsx** — The generated workbook with 8 sheets (1 data + 7 chart sheets, 28 charts total).
- **charts-line.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-line.py
# → charts-line.xlsx
```

## Chart Sheets

### Sheet: 1-Line Fundamentals

Four basic line charts covering every data input method and marker fundamentals.

```bash
# Inline named series with axis titles
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop series1="Product A:120,180,210,250" \
  --prop series2="Product B:90,140,160,200" \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop catTitle=Quarter --prop axisTitle=Revenue \
  --prop axisfont=9:58626E:Arial --prop gridlines=D9D9D9:0.5:dot

# Cell-range series (dotted syntax) with markers
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop series1.name=East \
  --prop series1.values=Sheet1!B2:B13 \
  --prop series1.categories=Sheet1!A2:A13 \
  --prop showMarkers=true --prop marker=circle:6:2E75B6 \
  --prop minorGridlines=EEEEEE:0.3:dot

# dataRange (auto-reads headers) with diamond markers
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop dataRange=Sheet1!A1:E13 \
  --prop showMarkers=true --prop marker=diamond:5:333333 \
  --prop legend=bottom --prop legendfont=9:58626E:Calibri

# Inline data shorthand with marker=none
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop 'data=Actual:80,120,160;Target:100,130,160' \
  --prop marker=none --prop legend=right
```

**Features:** `series1=Name:v1,v2`, `series1.name`/`.values`/`.categories` (cell range), `dataRange`, `data` (shorthand), `categories`, `colors`, `catTitle`, `axisTitle`, `axisfont`, `gridlines`, `minorGridlines`, `showMarkers`, `marker` (circle, diamond, none), `legend` (bottom, right), `legendfont`

### Sheet: 2-Line Styles

Four charts demonstrating visual styling — smoothing, dash patterns, markers, and transparency.

```bash
# Smooth curves with shadow, axes hidden
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop smooth=true --prop lineWidth=2.5 \
  --prop gridlines=none --prop axisVisible=false \
  --prop series.shadow=000000-4-315-2-40

# Dashed lines (applies to all series)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop lineDash=dash --prop lineWidth=2

# Marker styles with series outline
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop showMarkers=true --prop marker=square:7:4472C4 \
  --prop series.outline=FFFFFF-0.5

# Transparent lines on gradient plot area
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop lineWidth=3 --prop smooth=true \
  --prop transparency=30 \
  --prop plotFill=F0F4F8-D6E4F0:90 --prop chartFill=FFFFFF \
  --prop title.font=Georgia --prop title.size=14 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop roundedCorners=true
```

**Features:** `smooth`, `lineWidth`, `lineDash` (solid/dot/dash/dashdot/longdash/longdashdot/longdashdotdot), `marker` (square), `series.shadow` (color-blur-angle-dist-opacity), `series.outline`, `transparency`, `plotFill` (gradient), `chartFill`, `title.font`/`.size`/`.color`/`.bold`, `roundedCorners`, `gridlines=none`, `axisVisible=false`

### Sheet: 3-Line Variants

Four charts covering all line chart type variants.

```bash
# Stacked line — cumulative values
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=lineStacked \
  --prop majorTickMark=outside --prop tickLabelPos=low

# 100% stacked line — proportional
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=linePercentStacked \
  --prop axisNumFmt=0%

# 3D line with perspective
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line3d \
  --prop view3d=15,20,30 --prop style=3

# Stacked line with data table
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=lineStacked \
  --prop dataTable=true --prop legend=none
```

**Features:** `lineStacked`, `linePercentStacked`, `line3d`, `majorTickMark`, `tickLabelPos`, `axisNumFmt`, `view3d` (rotX,rotY,perspective), `style` (preset 1-48), `dataTable`, `legend=none`

### Sheet: 4-Axis & Gridlines

Four charts demonstrating every axis and gridline configuration.

```bash
# Custom axis scaling with axis lines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop axisMin=80 --prop axisMax=220 \
  --prop majorUnit=20 --prop minorUnit=10 \
  --prop axisLine=C00000:1.5:solid --prop catAxisLine=2E75B6:1.5:solid

# Logarithmic scale
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop logBase=10 \
  --prop marker=triangle:7:C00000

# Reversed value axis
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop axisReverse=true

# Display units with tick marks
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop dispUnits=thousands \
  --prop majorTickMark=outside --prop minorTickMark=inside \
  --prop marker=star:7:2E75B6
```

**Features:** `axisMin`, `axisMax`, `majorUnit`, `minorUnit`, `axisLine`, `catAxisLine`, `logBase` (logarithmic scale), `axisReverse` (flip direction), `dispUnits` (thousands/millions), `majorTickMark`, `minorTickMark`, `marker` (triangle, star)

### Sheet: 5-Labels & Legend

Four charts demonstrating data label and legend customization.

```bash
# Data labels with number format
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop dataLabels=true --prop labelPos=top \
  --prop labelFont=9:333333:true \
  --prop dataLabels.numFmt=#,##0 \
  --prop dataLabels.separator=": "

# Custom individual data labels (hide some, highlight peak with color + label)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop dataLabels=true \
  --prop dataLabel1.delete=true --prop dataLabel2.delete=true \
  --prop point4.color=C00000 \
  --prop dataLabel4.text="Peak: 210" --prop dataLabel4.y=0.15

# Legend overlay
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop legend=top --prop legend.overlay=true \
  --prop legendfont=10:1F4E79:Calibri

# Manual layout — plotArea, title, legend positioning
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop plotArea.x=0.12 --prop plotArea.y=0.18 \
  --prop plotArea.w=0.82 --prop plotArea.h=0.55 \
  --prop title.x=0.25 --prop title.y=0.02 \
  --prop legend.x=0.15 --prop legend.y=0.82 \
  --prop legend.w=0.7 --prop legend.h=0.12
```

**Features:** `dataLabels`, `labelPos` (top/center/insideEnd/outsideEnd/bestFit), `labelFont`, `dataLabels.numFmt`, `dataLabels.separator`, `dataLabel{N}.delete`, `dataLabel{N}.text`, `dataLabel{N}.y` (manual label position), `point{N}.color` (individual point color), `legend` (top), `legend.overlay`, `legendfont`, `plotArea.x/y/w/h`, `title.x/y`, `legend.x/y/w/h`

### Sheet: 6-Effects & Advanced

Four charts demonstrating advanced features — secondary axis, reference lines, effects, and conditional coloring.

```bash
# Secondary axis (dual scale)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop secondaryAxis=2 \
  --prop series1="Revenue:120,180,250,310" \
  --prop series2="Growth %:50,33,39,24"

# Reference line with longdash style
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop referenceLine=150:FF0000:1.5:dash \
  --prop lineDash=longdash

# Title glow/shadow effects
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop title.glow=4472C4-8-60 \
  --prop title.shadow=000000-3-315-2-40 \
  --prop series.shadow=000000-3-315-1-30

# Conditional coloring with chart/plot borders
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop colorRule=0:C00000:70AD47 \
  --prop referenceLine=0:888888:1:solid \
  --prop chartArea.border=D0D0D0:1:solid \
  --prop plotArea.border=E0E0E0:0.5:dot
```

**Features:** `secondaryAxis` (1-based series indices), `referenceLine` (value:color:width:dash), `title.glow` (color-radius-opacity), `title.shadow` (color-blur-angle-dist-opacity), `series.shadow`, `colorRule` (threshold:belowColor:aboveColor), `chartArea.border`, `plotArea.border`

### Sheet: 7-Line Elements

Four charts demonstrating line-chart-specific structural elements.

```bash
# Drop lines — vertical lines from points to X axis
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop dropLines=true

# High-low lines — connect highest and lowest series per category
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop hiLowLines=true

# Up-down bars with custom gain/loss colors
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop updownbars=100:70AD47:C00000

# 3D line with gap depth
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line3d \
  --prop gapDepth=300
```

**Features:** `dropLines` (vertical drop to axis), `hiLowLines` (high-low connectors), `updownbars` (gapWidth:upColor:downColor), `gapDepth` (3D depth spacing 0-500)

## Complete Feature Coverage

| Feature | Sheet |
|---------|-------|
| **Chart types:** line, lineStacked, linePercentStacked, line3d | 1, 3 |
| **Data input:** series, dataRange, data, series.name/values/categories | 1 |
| **Line styling:** smooth, lineWidth, lineDash, colors | 2 |
| **Markers:** circle, diamond, square, triangle, star, none, auto | 1, 2, 4 |
| **Axis scaling:** axisMin/Max, majorUnit, minorUnit | 4 |
| **Axis features:** logBase, axisReverse, dispUnits, axisNumFmt | 3, 4 |
| **Axis lines:** axisLine, catAxisLine | 4 |
| **Axis visibility:** axisVisible | 2 |
| **Tick marks:** majorTickMark, minorTickMark, tickLabelPos | 3, 4 |
| **Gridlines:** gridlines, minorGridlines, gridlines=none | 1, 2, 4 |
| **Data labels:** dataLabels, labelPos, labelFont, numFmt, separator | 5 |
| **Custom labels:** dataLabel{N}.text, dataLabel{N}.delete, dataLabel{N}.y | 5 |
| **Point color:** point{N}.color | 5 |
| **Legend:** position, legendfont, legend.overlay, legend=none | 1, 3, 5 |
| **Layout:** plotArea.x/y/w/h, title.x/y, legend.x/y/w/h | 5 |
| **Effects:** series.shadow, series.outline, transparency | 2, 6 |
| **Title styling:** font, size, color, bold, glow, shadow | 2, 6 |
| **Fills:** plotFill, chartFill (solid + gradient) | 2, 3, 6 |
| **Borders:** chartArea.border, plotArea.border | 6 |
| **Advanced:** secondaryAxis, referenceLine, colorRule | 6 |
| **Line elements:** dropLines, hiLowLines, upDownBars | 7 |
| **3D:** view3d, gapDepth, style | 3, 7 |
| **Other:** dataTable, roundedCorners | 2, 3 |

## Inspect the Generated File

```bash
officecli query charts-line.xlsx chart
officecli get charts-line.xlsx "/1-Line Fundamentals/chart[1]"
```
