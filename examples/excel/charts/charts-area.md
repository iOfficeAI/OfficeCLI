# Area Charts Showcase

This demo consists of three files that work together:

- **charts-area.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-area.xlsx** — The generated workbook with 6 sheets (1 data + 5 chart sheets, 20 charts total).
- **charts-area.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-area.py
# → charts-area.xlsx
```

## Chart Sheets

### Sheet: 1-Area Fundamentals

Four area charts covering data input methods, transparency, area fills, and gradients.

```bash
# Basic area with dataRange and axis titles
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop dataRange=Sheet1!A1:E13 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop catTitle=Month --prop axisTitle=Visitors \
  --prop gridlines=D9D9D9:0.5:dot

# Inline series with transparency
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop series1="Subscriptions:120,180,210,250" \
  --prop series2="One-time:90,140,160,200" \
  --prop transparency=40 --prop legend=bottom

# Area with areafill gradient (single series)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop series1="Users:3200,3800,4500,5100,5800,6400" \
  --prop areafill=4472C4-BDD7EE:90 --prop legend=none

# Per-series gradient fills
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90' \
  --prop legend=right --prop legendfont=10:333333:Calibri
```

**Features:** `area`, `dataRange`, `categories`, `colors`, `catTitle`, `axisTitle`, `gridlines`, `transparency`, `areafill` (gradient from-to:angle), `gradients` (per-series), `legend` (bottom, right, none), `legendfont`

### Sheet: 2-Area Variants

Four charts covering all area chart type variants — stacked, percent stacked, and 3D.

```bash
# Stacked area with solid plot fill
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=areaStacked \
  --prop plotFill=F5F5F5 --prop roundedCorners=true

# 100% stacked area with axis number format
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=areaPercentStacked \
  --prop axisNumFmt=0% --prop axisLine=333333:1:solid

# 3D area with perspective rotation
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area3d \
  --prop view3d=20,25,15

# 3D area with multiple series and gridlines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area3d \
  --prop view3d=15,20,20 --prop gridlines=D9D9D9:0.5:dot
```

**Features:** `areaStacked`, `areaPercentStacked`, `area3d`, `plotFill` (solid), `roundedCorners`, `axisNumFmt`, `axisLine`, `view3d` (rotX,rotY,perspective)

### Sheet: 3-Area Styling

Four charts demonstrating visual styling — title effects, shadows, gridlines, and fills.

```bash
# Title styling with shadow
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop title.shadow=000000-3-315-2-30

# Series shadow, outline, and smooth curve
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop smooth=true \
  --prop series.shadow=000000-4-315-2-40 \
  --prop series.outline=333333-1

# Axis font with gridlines and minor gridlines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop axisfont=9:58626E:Arial \
  --prop gridlines=D9D9D9:0.5:dot \
  --prop minorGridlines=EEEEEE:0.3:dot

# Chart fill, plot fill gradient, and borders
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop chartFill=FAFAFA \
  --prop plotFill=E8F0FE-D6E4F0:90 \
  --prop chartArea.border=D0D0D0:1:solid \
  --prop plotArea.border=E0E0E0:0.5:dot
```

**Features:** `title.font`/`.size`/`.color`/`.bold`/`.shadow`, `smooth`, `series.shadow` (color-blur-angle-dist-opacity), `series.outline` (color-width), `axisfont` (size:color:font), `gridlines`, `minorGridlines`, `chartFill`, `plotFill` (gradient), `chartArea.border`, `plotArea.border`, `roundedCorners`

### Sheet: 4-Labels & Legend

Four charts demonstrating data label and legend customization plus manual layout.

```bash
# Data labels with position, font, and number format
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop dataLabels=true --prop labelPos=top \
  --prop labelFont=9:333333:true \
  --prop dataLabels.numFmt=#,##0

# Individual label deletion and per-point colors
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop dataLabels=true \
  --prop dataLabel1.delete=true --prop dataLabel2.delete=true \
  --prop point4.color=C00000

# Legend overlay with font styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop legend=right --prop legendfont=10:1F4E79:Calibri \
  --prop legend.overlay=true

# Manual layout — plotArea, title, legend positioning
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop plotArea.x=0.12 --prop plotArea.y=0.18 \
  --prop plotArea.w=0.82 --prop plotArea.h=0.55 \
  --prop title.x=0.25 --prop title.y=0.02 \
  --prop legend.x=0.15 --prop legend.y=0.82 \
  --prop legend.w=0.7 --prop legend.h=0.12
```

**Features:** `dataLabels`, `labelPos` (top), `labelFont`, `dataLabels.numFmt`, `dataLabel{N}.delete`, `point{N}.color`, `legend` (right), `legendfont`, `legend.overlay`, `plotArea.x/y/w/h`, `title.x/y`, `legend.x/y/w/h`

### Sheet: 5-Advanced

Four charts demonstrating advanced features — secondary axis, reference lines, axis scaling, and effects.

```bash
# Secondary axis (dual scale)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop secondaryAxis=2 \
  --prop series1="Revenue:120,180,250,310,280,340" \
  --prop series2="Conv %:2.1,2.8,3.2,3.9,3.5,4.1"

# Reference line (target/threshold)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop referenceLine=100:FF0000:1.5:dash \
  --prop areafill=4472C4-BDD7EE:90

# Axis scaling with display units
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop axisMin=3000 --prop axisMax=7000 \
  --prop majorUnit=500 --prop dispUnits=thousands

# Color rule with title glow and series shadow
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop colorRule=50:C00000:70AD47 \
  --prop referenceLine=50:888888:1:solid \
  --prop title.glow=4472C4-8-60 \
  --prop series.shadow=000000-3-315-1-30
```

**Features:** `secondaryAxis` (1-based series index), `referenceLine` (value:color:width:dash), `axisMin`, `axisMax`, `majorUnit`, `dispUnits` (thousands), `colorRule` (threshold:belowColor:aboveColor), `title.glow` (color-radius-opacity), `areafill`

## Complete Feature Coverage

| Feature | Sheet |
|---------|-------|
| **Chart types:** area, areaStacked, areaPercentStacked, area3d | 1, 2 |
| **Data input:** dataRange, series, categories, colors | 1 |
| **Area fills:** areafill (gradient), gradients (per-series), transparency | 1, 5 |
| **Axis titles:** catTitle, axisTitle | 1, 3 |
| **Axis scaling:** axisMin/Max, majorUnit, dispUnits | 5 |
| **Axis features:** axisNumFmt, axisLine | 2 |
| **Gridlines:** gridlines, minorGridlines | 1, 3 |
| **Data labels:** dataLabels, labelPos, labelFont, numFmt | 4 |
| **Custom labels:** dataLabel{N}.delete | 4 |
| **Point color:** point{N}.color | 4 |
| **Legend:** position, legendfont, legend.overlay, legend=none | 1, 4 |
| **Layout:** plotArea.x/y/w/h, title.x/y, legend.x/y/w/h | 4 |
| **Effects:** series.shadow, series.outline, smooth | 3 |
| **Title styling:** font, size, color, bold, shadow, glow | 3, 5 |
| **Fills:** plotFill, chartFill (solid + gradient) | 2, 3 |
| **Borders:** chartArea.border, plotArea.border | 3 |
| **Advanced:** secondaryAxis, referenceLine, colorRule | 5 |
| **3D:** view3d | 2 |
| **Other:** roundedCorners | 2, 3 |

## Inspect the Generated File

```bash
officecli query charts-area.xlsx chart
officecli get charts-area.xlsx "/1-Area Fundamentals/chart[1]"
```
