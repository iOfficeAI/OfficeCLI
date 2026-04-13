# Waterfall Charts Showcase

This demo consists of three files that work together:

- **charts-waterfall.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-waterfall.xlsx** — The generated workbook with 5 sheets (1 default + 4 chart sheets, 16 charts total).
- **charts-waterfall.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-waterfall.py
# → charts-waterfall.xlsx
```

## Chart Sheets

### Sheet: 1-Waterfall Fundamentals

Four waterfall chart variants covering basic P&L, budget analysis, quarterly flow, and title styling.

```bash
# Basic P&L waterfall with increase/decrease/total colors
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop data="Start:1000,Revenue:500,Costs:-300,Tax:-100,Net:1100" \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=FF0000 \
  --prop totalColor=4472C4 \
  --prop dataLabels=true

# Budget waterfall with blue/red/amber theme
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop data="Budget:5000,Sales:2000,Marketing:-800,Ops:-600,Net:5600" \
  --prop increaseColor=2E75B6 \
  --prop decreaseColor=C00000 \
  --prop totalColor=FFC000 \
  --prop legend=bottom

# Quarterly cash flow with 10 data points
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop data="Opening:3000,Q1 Sales:1200,Q1 Costs:-500,...,Closing:6000" \
  --prop increaseColor=70AD47 --prop decreaseColor=ED7D31 --prop totalColor=4472C4

# Waterfall with styled title
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true
```

**Features:** `chartType=waterfall`, `data=` name:value pairs (positive=increase, negative=decrease), `increaseColor`, `decreaseColor`, `totalColor`, `dataLabels`, `legend=bottom`, `title.font/size/color/bold`

### Sheet: 2-Waterfall Styling

Four waterfall charts demonstrating visual styling options.

```bash
# Title with font, size, color, bold, and shadow
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop title.font=Trebuchet MS --prop title.size=18 \
  --prop title.color=833C0B --prop title.bold=true \
  --prop title.shadow=000000-3-315-2-30

# Series shadow, plot/chart fills, rounded corners
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop series.shadow=000000-4-315-2-30 \
  --prop plotFill=F0F0F0 --prop chartFill=FAFAFA \
  --prop roundedCorners=true

# Gridline color and axis font
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop gridlineColor=CCCCCC \
  --prop axisfont=10:333333:Calibri

# Chart area and plot area borders
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop chartArea.border=4472C4-2 \
  --prop plotArea.border=A5A5A5-1
```

**Features:** `title.shadow`, `series.shadow`, `plotFill`, `chartFill`, `roundedCorners`, `gridlineColor`, `axisfont`, `chartArea.border`, `plotArea.border`

### Sheet: 3-Waterfall Labels & Axis

Four waterfall charts demonstrating data labels, axis configuration, and layout control.

```bash
# Data labels with font and number format
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop dataLabels=true \
  --prop labelFont=10:333333:true \
  --prop dataLabels.numFmt=#,##0

# Custom axis range and tick interval
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop axisMin=0 --prop axisMax=3500 --prop majorUnit=500

# Legend position and font
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop legend=right \
  --prop legendfont=10:1F4E79:Helvetica

# Manual plot area layout
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop plotArea.x=0.15 --prop plotArea.y=0.15 \
  --prop plotArea.w=0.75 --prop plotArea.h=0.70
```

**Features:** `dataLabels`, `labelFont`, `dataLabels.numFmt`, `axisMin`, `axisMax`, `majorUnit`, `legend=right`, `legendfont`, `plotArea.x/y/w/h`

### Sheet: 4-Waterfall Advanced

Four waterfall charts demonstrating advanced features and large datasets.

```bash
# Reference line overlay
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop referenceLine=2000:Target-FF0000-dash-2

# Value axis and category axis line styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop axisLine=333333-2 \
  --prop catAxisLine=333333-2

# Title glow and shadow effects
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop title.glow=4472C4-8 \
  --prop title.shadow=000000-3-315-2-30

# Large dataset (12 categories) with small axis font
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop data="Revenue:8500,COGS:-3400,...,Net Income:1050" \
  --prop dataLabels=true \
  --prop axisfont=8:333333:Calibri
```

**Features:** `referenceLine`, `axisLine`, `catAxisLine`, `title.glow`, `title.shadow`, large dataset (12 categories)

## Property Coverage

| Property | Sheet |
|---|---|
| `chartType=waterfall` | 1, 2, 3, 4 |
| `data=` (name:value pairs) | 1, 2, 3, 4 |
| `increaseColor` | 1, 2, 3, 4 |
| `decreaseColor` | 1, 2, 3, 4 |
| `totalColor` | 1, 2, 3, 4 |
| `dataLabels` | 1, 3, 4 |
| `legend` | 1, 3 |
| `title.font/size/color/bold` | 1, 2 |
| `title.shadow` | 2, 4 |
| `title.glow` | 4 |
| `series.shadow` | 2 |
| `plotFill`, `chartFill` | 2 |
| `roundedCorners` | 2 |
| `gridlineColor` | 2 |
| `axisfont` | 2, 4 |
| `chartArea.border` | 2 |
| `plotArea.border` | 2 |
| `labelFont` | 3 |
| `dataLabels.numFmt` | 3 |
| `axisMin/Max`, `majorUnit` | 3 |
| `legendfont` | 3 |
| `plotArea.x/y/w/h` | 3 |
| `referenceLine` | 4 |
| `axisLine`, `catAxisLine` | 4 |

## Inspect the Generated File

```bash
officecli query charts-waterfall.xlsx chart
officecli get charts-waterfall.xlsx "/1-Waterfall Fundamentals/chart[1]"
```
