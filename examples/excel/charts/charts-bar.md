# Bar (Horizontal) Charts Showcase

This demo consists of three files that work together:

- **charts-bar.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-bar.xlsx** — The generated workbook with 7 sheets (1 data + 6 chart sheets, 24 charts total).
- **charts-bar.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-bar.py
# → charts-bar.xlsx
```

## Chart Sheets

### Sheet: 1-Bar Fundamentals

Four basic horizontal bar charts covering data input variants, colors, stacking, and shorthand syntax.

```bash
# Basic bar from cell range with axis titles and gridlines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop dataRange=Sheet1!A1:B9 \
  --prop catTitle=Department --prop axisTitle=Score \
  --prop axisfont=9:333333:Arial \
  --prop gridlines=D9D9D9:0.5:dot

# Inline series with custom colors and data labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop series1="Satisfaction:85,72,91,68,78" \
  --prop colors=4472C4,ED7D31,70AD47,FFC000,5B9BD5 \
  --prop gapwidth=80 --prop dataLabels=outsideEnd

# Stacked bar with series outline
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=barStacked \
  --prop series1="Q1:30,18,25,12" --prop series2="Q2:35,20,28,14" \
  --prop overlap=0 --prop series.outline=FFFFFF-0.5

# data= shorthand with legend at bottom
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop 'data=Technical:45,38,52;Soft Skills:20,28,18;Compliance:12,15,10' \
  --prop legend=bottom
```

**Features:** `bar`, `barStacked`, `dataRange`, `catTitle`, `axisTitle`, `axisfont`, `gridlines`, `colors`, `gapwidth`, `dataLabels=outsideEnd`, `overlap`, `series.outline`, `data=` shorthand, `legend=bottom`

### Sheet: 2-Bar Variants

Four bar chart type variants: stacked, 100% stacked, 3D, and 3D cylinder.

```bash
# Stacked bar with tight gap
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=barStacked \
  --prop gapwidth=50

# 100% stacked with percentage axis and reference line
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=barPercentStacked \
  --prop axisNumFmt=0% \
  --prop referenceLine=0.5:FF0000:Target:dash

# 3D bar with perspective
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar3d \
  --prop view3d=10,30,20 --prop style=3

# 3D bar with cylinder shape
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar3d \
  --prop shape=cylinder --prop gapwidth=60
```

**Features:** `barStacked`, `barPercentStacked`, `bar3d`, `gapwidth`, `axisNumFmt=0%`, `referenceLine` (with label and dash), `view3d`, `style`, `shape=cylinder`

### Sheet: 3-Bar Styling

Four charts demonstrating visual styling: title formatting, shadows, gradients, and background fills.

```bash
# Title font, size, color, bold
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true

# Series shadow and outline
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop series.shadow=000000-4-315-2-30 \
  --prop series.outline=1F4E79-1

# Per-bar gradient fills (angle=0 for horizontal)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop 'gradients=1F4E79-5B9BD5:0;C55A11-F4B183:0;...' \
  --prop labelFont=9:333333:true

# Plot/chart fill with transparency and rounded corners
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop plotFill=F0F4F8-D6E4F0:90 --prop chartFill=FFFFFF \
  --prop transparency=20 --prop roundedCorners=true
```

**Features:** `title.font/size/color/bold`, `series.shadow`, `series.outline`, `gradients` (per-bar), `labelFont`, `plotFill` gradient, `chartFill`, `transparency`, `roundedCorners`

### Sheet: 4-Axis & Labels

Four charts exploring axis configuration and data label customization.

```bash
# Custom axis scale with gridlines styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop axisMin=50 --prop axisMax=250 --prop majorUnit=50 \
  --prop gridlines=D0D0D0:0.5:solid \
  --prop minorGridlines=EEEEEE:0.3:dot \
  --prop axisLine=C00000:1.5:solid --prop catAxisLine=2E75B6:1.5:solid

# Log scale, reversed axis, display units
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop logBase=10 --prop axisReverse=true \
  --prop dispUnits=thousands

# Data labels with font, number format, separator
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop labelFont=10:1F4E79:true \
  --prop dataLabels.numFmt=#,##0 --prop "dataLabels.separator=: "

# Per-point label delete/text and per-point color (highlight winner)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop dataLabel1.delete=true --prop dataLabel4.text="Winner!" \
  --prop point4.color=C00000 --prop point2.color=2E75B6
```

**Features:** `axisMin`, `axisMax`, `majorUnit`, `gridlines`, `minorGridlines`, `axisLine`, `catAxisLine`, `logBase`, `axisReverse`, `dispUnits`, `dataLabels`, `labelPos`, `labelFont`, `dataLabels.numFmt`, `dataLabels.separator`, `dataLabel{N}.delete`, `dataLabel{N}.text`, `point{N}.color`

### Sheet: 5-Legend & Layout

Four charts covering legend configuration, manual layout, and dual-axis support.

```bash
# Legend on right side
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop legend=right

# Legend font styling with overlay
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop legend=top --prop legend.overlay=true \
  --prop legendfont=10:1F4E79:Calibri

# Manual layout: plotArea, title, and legend positioning
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop plotArea.x=0.25 --prop plotArea.y=0.15 \
  --prop plotArea.w=0.70 --prop plotArea.h=0.60 \
  --prop title.x=0.20 --prop title.y=0.02 \
  --prop legend.x=0.25 --prop legend.y=0.82

# Secondary axis with chart/plot area borders
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop secondaryAxis=2 \
  --prop chartArea.border=D0D0D0:1:solid \
  --prop plotArea.border=E0E0E0:0.5:dot
```

**Features:** `legend=right/top/bottom`, `legend.overlay`, `legendfont`, `plotArea.x/y/w/h`, `title.x/y`, `legend.x/y/w/h`, `secondaryAxis`, `chartArea.border`, `plotArea.border`

### Sheet: 6-Advanced

Four charts with advanced features: reference lines, conditional coloring, effects, and data tables.

```bash
# Reference line with label
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop referenceLine=79:FF0000:Average:dash

# Conditional coloring (profit/loss)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop colorRule=0:C00000:70AD47 \
  --prop referenceLine=0:888888:1:solid

# Title glow, title shadow, series shadow
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop title.glow=4472C4-8-60 \
  --prop title.shadow=000000-3-315-2-40 \
  --prop series.shadow=000000-3-315-1-30

# Error bars and data table
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop errBars=percent:10 --prop dataTable=true \
  --prop legend=none
```

**Features:** `referenceLine` (with label), `colorRule` (threshold coloring), `title.glow`, `title.shadow`, `series.shadow`, `errBars=percent:10`, `dataTable=true`

## Feature Coverage

| Feature | Sheet |
|---|---|
| `bar` (basic horizontal) | 1, 3, 4, 5, 6 |
| `barStacked` | 1, 2 |
| `barPercentStacked` | 2 |
| `bar3d` | 2 |
| `bar3d shape=cylinder` | 2 |
| `dataRange` (cell reference) | 1, 3, 5, 6 |
| `data=` shorthand | 1 |
| `series1=Name:values` | 1, 2, 3, 4, 5, 6 |
| `colors` | 1, 2, 3, 4, 5, 6 |
| `gapwidth` | 1, 2, 4, 6 |
| `overlap` | 1 |
| `dataLabels` / `labelPos` | 1, 3, 4, 6 |
| `labelFont` | 3, 4, 6 |
| `dataLabels.numFmt` | 4 |
| `dataLabels.separator` | 4 |
| `dataLabel{N}.delete/text` | 4 |
| `point{N}.color` | 4 |
| `catTitle` / `axisTitle` | 1 |
| `axisfont` | 1 |
| `axisMin/Max` / `majorUnit` | 4 |
| `gridlines` / `minorGridlines` | 1, 4, 6 |
| `axisLine` / `catAxisLine` | 4 |
| `logBase` | 4 |
| `axisReverse` | 4 |
| `dispUnits` | 4 |
| `axisNumFmt` | 2 |
| `legend` positions | 1, 2, 5, 6 |
| `legendfont` | 5 |
| `legend.overlay` | 5 |
| `title.font/size/color/bold` | 3 |
| `title.glow` / `title.shadow` | 6 |
| `series.shadow` | 3, 6 |
| `series.outline` | 1, 3 |
| `gradients` | 3 |
| `plotFill` / `chartFill` | 3, 6 |
| `transparency` | 3 |
| `roundedCorners` | 3 |
| `referenceLine` | 2, 6 |
| `colorRule` | 6 |
| `secondaryAxis` | 5 |
| `chartArea.border` / `plotArea.border` | 5 |
| `plotArea.x/y/w/h` | 5 |
| `title.x/y` | 5 |
| `legend.x/y/w/h` | 5 |
| `view3d` / `style` | 2 |
| `shape=cylinder` | 2 |
| `errBars` | 6 |
| `dataTable` | 6 |

## Inspect the Generated File

```bash
officecli query charts-bar.xlsx chart
officecli get charts-bar.xlsx "/1-Bar Fundamentals/chart[1]"
```
