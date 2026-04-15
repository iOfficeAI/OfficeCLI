# Column Charts Showcase

This demo consists of three files that work together:

- **charts-column.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-column.xlsx** — The generated workbook with 8 sheets (1 data + 7 chart sheets, 28 charts total).
- **charts-column.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-column.py
# → charts-column.xlsx
```

## Chart Sheets

### Sheet: 1-Column Fundamentals

Four basic column charts covering every data input method.

```bash
# dataRange with axis titles and axis font
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop dataRange=Sheet1!A1:E13 \
  --prop catTitle=Month --prop axisTitle=Revenue \
  --prop axisfont=9:58626E:Arial --prop gridlines=D9D9D9:0.5:dot

# Inline named series with gap width
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop series1="Laptops:320,280,350,310" \
  --prop series2="Phones:450,420,480,460" \
  --prop categories=Jan,Feb,Mar,Apr \
  --prop colors=2E75B6,C00000,70AD47 \
  --prop gapwidth=80

# Cell-range series (dotted syntax)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop series1.name=East \
  --prop series1.values=Sheet1!B2:B13 \
  --prop series1.categories=Sheet1!A2:A13 \
  --prop minorGridlines=EEEEEE:0.3:dot

# Inline data shorthand
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop 'data=Team A:85,92,78;Team B:70,80,85' \
  --prop categories=Mon,Tue,Wed \
  --prop legend=right
```

**Features:** `series1=Name:v1,v2`, `series1.name`/`.values`/`.categories` (cell range), `dataRange`, `data` (shorthand), `categories`, `colors`, `catTitle`, `axisTitle`, `axisfont`, `gridlines`, `minorGridlines`, `gapwidth`, `legend` (bottom, right)

### Sheet: 2-Column Variants

Four charts covering all column chart type variants.

```bash
# Stacked column with center labels and series outline
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=columnStacked \
  --prop dataLabels=center \
  --prop series.outline=FFFFFF-0.5

# 100% stacked column — proportional
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=columnPercentStacked \
  --prop axisNumFmt=0%

# 3D column with perspective
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column3d \
  --prop view3d=15,20,30 --prop style=3

# 3D column with gap depth
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column3d \
  --prop gapDepth=200
```

**Features:** `columnStacked`, `columnPercentStacked`, `column3d`, `dataLabels=center`, `series.outline`, `axisNumFmt`, `view3d` (rotX,rotY,perspective), `style` (preset 1-48), `gapDepth`

### Sheet: 3-Column Styling

Four charts demonstrating visual styling — title formatting, shadows, gradients, and transparency.

```bash
# Styled title
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true

# Series shadow and outline effects
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop series.shadow=000000-4-315-2-40 \
  --prop series.outline=FFFFFF-0.5

# Per-series gradient fills
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90;70AD47-C5E0B4:90'

# Transparent columns on gradient background
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop transparency=30 \
  --prop plotFill=F0F4F8-D6E4F0:90 --prop chartFill=FFFFFF \
  --prop roundedCorners=true
```

**Features:** `title.font`/`.size`/`.color`/`.bold`, `series.shadow` (color-blur-angle-dist-opacity), `series.outline`, `gradients` (per-series), `transparency`, `plotFill` (gradient), `chartFill`, `roundedCorners`

### Sheet: 4-Axis & Gridlines

Four charts demonstrating every axis and gridline configuration.

```bash
# Custom axis scaling with axis lines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop axisMin=50 --prop axisMax=250 \
  --prop majorUnit=50 --prop minorUnit=25 \
  --prop axisLine=C00000:1.5:solid --prop catAxisLine=2E75B6:1.5:solid

# Logarithmic scale with reversed axis
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop logBase=10 --prop axisReverse=true

# Display units with tick marks
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop dispUnits=thousands --prop axisNumFmt=#,##0 \
  --prop majorTickMark=outside --prop minorTickMark=inside

# Hidden axes with data table
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop gridlines=none --prop axisVisible=false \
  --prop dataTable=true --prop legend=none
```

**Features:** `axisMin`, `axisMax`, `majorUnit`, `minorUnit`, `axisLine`, `catAxisLine`, `logBase` (logarithmic scale), `axisReverse` (flip direction), `dispUnits` (thousands/millions), `axisNumFmt`, `majorTickMark`, `minorTickMark`, `axisVisible`, `dataTable`, `gridlines=none`, `legend=none`

### Sheet: 5-Labels & Legend

Four charts demonstrating data label and legend customization.

```bash
# Data labels with number format
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop labelFont=9:333333:true \
  --prop dataLabels.numFmt=#,##0

# Custom individual labels (hide some, highlight peak)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop dataLabels=true \
  --prop dataLabel1.delete=true --prop dataLabel2.delete=true \
  --prop point4.color=C00000 --prop dataLabel4.text=Peak!

# Legend overlay with styled font
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop legend=right --prop legend.overlay=true \
  --prop legendfont=10:333333:Calibri --prop plotFill=F5F5F5

# Manual layout — plotArea, title, legend positioning
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop plotArea.x=0.12 --prop plotArea.y=0.18 \
  --prop plotArea.w=0.82 --prop plotArea.h=0.55 \
  --prop title.x=0.25 --prop title.y=0.02 \
  --prop legend.x=0.15 --prop legend.y=0.82 \
  --prop legend.w=0.7 --prop legend.h=0.12
```

**Features:** `dataLabels`, `labelPos` (outsideEnd/center/insideEnd/insideBase), `labelFont`, `dataLabels.numFmt`, `dataLabel{N}.delete`, `dataLabel{N}.text`, `point{N}.color`, `legend` (right), `legend.overlay`, `legendfont`, `plotFill`, `plotArea.x/y/w/h`, `title.x/y`, `legend.x/y/w/h`

### Sheet: 6-Effects & Advanced

Four charts demonstrating advanced features — secondary axis, reference lines, effects, and conditional coloring.

```bash
# Secondary axis (dual scale)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop secondaryAxis=2 \
  --prop series1="Revenue:120,180,250,310" \
  --prop series2="Growth %:50,33,39,24"

# Reference line (target threshold)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop referenceLine=150:FF0000:1.5:dash

# Title glow/shadow effects
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop title.glow=4472C4-8-60 \
  --prop title.shadow=000000-3-315-2-40 \
  --prop series.shadow=000000-3-315-1-30

# Conditional coloring with chart/plot borders
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop colorRule=0:C00000:70AD47 \
  --prop referenceLine=0:888888:1:solid \
  --prop chartArea.border=D0D0D0:1:solid \
  --prop plotArea.border=E0E0E0:0.5:dot
```

**Features:** `secondaryAxis` (1-based series indices), `referenceLine` (value:color:width:dash), `title.glow` (color-radius-opacity), `title.shadow` (color-blur-angle-dist-opacity), `series.shadow`, `colorRule` (threshold:belowColor:aboveColor), `chartArea.border`, `plotArea.border`

### Sheet: 7-Bar Shape & Gap

Four charts demonstrating column gap width, overlap, and 3D bar shapes.

```bash
# Narrow gap (bars close together)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop gapwidth=30

# Wide gap with negative overlap (separated bars within group)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop gapwidth=200 --prop overlap=-50

# Cylinder shape (3D)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column3d \
  --prop shape=cylinder --prop view3d=15,20,30

# Cone shape (3D)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column3d \
  --prop shape=cone --prop view3d=15,20,30
```

**Features:** `gapwidth` (0-500), `overlap` (-100 to 100, negative = separated), `shape` (cylinder, cone, pyramid — 3D column shapes)

## Complete Feature Coverage

| Feature | Sheet |
|---------|-------|
| **Chart types:** column, columnStacked, columnPercentStacked, column3d | 1, 2 |
| **Data input:** series, dataRange, data, series.name/values/categories | 1 |
| **Colors:** colors, gradients | 1, 3 |
| **Gap & overlap:** gapwidth, overlap | 1, 7 |
| **Axis scaling:** axisMin/Max, majorUnit, minorUnit | 4 |
| **Axis features:** logBase, axisReverse, dispUnits, axisNumFmt | 2, 4 |
| **Axis lines:** axisLine, catAxisLine | 4 |
| **Axis visibility:** axisVisible | 4 |
| **Tick marks:** majorTickMark, minorTickMark | 4 |
| **Gridlines:** gridlines, minorGridlines, gridlines=none | 1, 4 |
| **Data labels:** dataLabels, labelPos, labelFont, numFmt | 2, 5 |
| **Custom labels:** dataLabel{N}.text, dataLabel{N}.delete | 5 |
| **Point color:** point{N}.color | 5 |
| **Legend:** position, legendfont, legend.overlay, legend=none | 1, 4, 5 |
| **Layout:** plotArea.x/y/w/h, title.x/y, legend.x/y/w/h | 5 |
| **Effects:** series.shadow, series.outline, transparency | 2, 3 |
| **Title styling:** font, size, color, bold, glow, shadow | 3, 6 |
| **Fills:** plotFill, chartFill (solid + gradient) | 3, 5 |
| **Borders:** chartArea.border, plotArea.border | 6 |
| **Advanced:** secondaryAxis, referenceLine, colorRule | 6 |
| **3D:** view3d, gapDepth, style, shape (cylinder/cone/pyramid) | 2, 7 |
| **Other:** dataTable, roundedCorners, catTitle, axisTitle, axisfont | 1, 3, 4 |

## Inspect the Generated File

```bash
officecli query charts-column.xlsx chart
officecli get charts-column.xlsx "/1-Column Fundamentals/chart[1]"
```
