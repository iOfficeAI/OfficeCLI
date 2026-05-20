# Basic Charts Showcase

This demo consists of three files that work together:

- **charts-basic.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments, then executed by the script.
- **charts-basic.xlsx** — The generated workbook with 8 sheets (1 data + 7 chart sheets, 28 charts total). Open in Excel to see the rendered charts.
- **charts-basic.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-basic.py
# → charts-basic.xlsx
```

## Source Data

**Sheet1**: 12 months of regional sales data (East, South, North, West) used by all charts.

## Chart Sheets

### Sheet: 1-Column Charts

Four column chart variants demonstrating the column family.

```bash
# Basic clustered column with axis titles and axis font
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column \
  --prop title="Regional Sales" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop catTitle=Month --prop axisTitle=Sales \
  --prop axisfont=9:58626E:Arial \
  --prop gridlines=D9D9D9:0.5:dot

# Stacked column with custom colors, data labels, gap control, series outline
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=columnStacked \
  --prop colors=2E75B6,70AD47,FFC000,C00000 \
  --prop dataLabels=true --prop labelPos=center \
  --prop gapwidth=60 \
  --prop series.outline=FFFFFF-0.5

# 100% stacked with legend positioning and plot fill
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=columnPercentStacked \
  --prop legend=bottom --prop legendfont=9:8B949E \
  --prop plotFill=F5F5F5

# 3D column with perspective and title styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=column3d \
  --prop view3d=15,20,30 \
  --prop title.font=Calibri --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true
```

**Features:** `column`, `columnStacked`, `columnPercentStacked`, `column3d`, `dataRange`, `catTitle`, `axisTitle`, `axisfont`, `gridlines`, `colors`, `dataLabels`, `labelPos`, `gapwidth`, `series.outline`, `legend`, `legendfont`, `plotFill`, `view3d`, `title.font/size/color/bold`

### Sheet: 2-Bar Charts

Four horizontal bar chart variants.

```bash
# Horizontal bar with inline data and gap control
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar \
  --prop 'data=East:198;South:158;North:142;West:180' \
  --prop gapwidth=80 \
  --prop dataLabels=true --prop labelPos=outsideEnd

# Stacked bar with named series and overlap
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=barStacked \
  --prop series1=H1:663,598,528,661 \
  --prop series2=H2:833,718,669,868 \
  --prop gapwidth=50 --prop overlap=0

# 100% stacked bar with reference line and axis lines
# Note: value axis of a barPercentStacked chart is 0-1 (= 0%-100%), so a 50% line = 0.5
# referenceLine forms: value | value:color | value:color:label | value:color:width:dash
#                      | value:color:label:dash | value:color:width:dash:label
# Width is in points (default 1.5pt). e.g. 0.5:FF0000:2:dash draws a 2pt dashed line.
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=barPercentStacked \
  --prop referenceLine=0.5:FF0000:Target:dash \
  --prop axisLine=333333:1:solid \
  --prop catAxisLine=333333:1:solid

# 3D bar with chart area fill and preset style
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bar3d \
  --prop view3d=10,30,20 \
  --prop chartFill=F2F2F2 \
  --prop style=3
```

**Features:** `bar`, `barStacked`, `barPercentStacked`, `bar3d`, inline `data`, named `series`, `gapwidth`, `overlap`, `labelPos=outsideEnd`, `referenceLine`, `axisLine`, `catAxisLine`, `chartFill`, `style`

### Sheet: 3-Line Charts

Four line chart variants with markers, smoothing, and data tables.

```bash
# Line with cell-range series (dotted syntax) and markers
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop series1.name=East \
  --prop series1.values=Sheet1!B2:B13 \
  --prop series1.categories=Sheet1!A2:A13 \
  --prop showMarkers=true --prop marker=circle:6:2E75B6 \
  --prop gridlines=D9D9D9:0.5:dot \
  --prop minorGridlines=EEEEEE:0.3:dot

# Smooth line with series shadow
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop smooth=true --prop lineWidth=2.5 \
  --prop gridlines=none \
  --prop series.shadow=000000-4-315-2-40

# Stacked line with tick marks
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=lineStacked \
  --prop majorTickMark=outside --prop tickLabelPos=low

# Dashed line with data table and hidden legend
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=line \
  --prop lineDash=dash --prop lineWidth=1.5 \
  --prop dataTable=true --prop legend=none
```

**Features:** `series1.name/values/categories` (cell range), `showMarkers`, `marker` (style:size:color), `smooth`, `lineWidth`, `lineDash`, `gridlines`, `minorGridlines`, `series.shadow`, `lineStacked`, `majorTickMark`, `tickLabelPos`, `dataTable`, `legend=none`

### Sheet: 4-Area Charts

Four area chart variants with transparency and gradients.

```bash
# Area with transparency and gradient
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area \
  --prop transparency=40 \
  --prop gradient=4472C4-BDD7EE:90

# Stacked area with plot fill and rounded corners
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=areaStacked \
  --prop plotFill=F5F5F5 --prop roundedCorners=true

# 100% stacked area with axis visibility control
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=areaPercentStacked \
  --prop axisVisible=true --prop axisLine=999999:0.5:solid

# 3D area with perspective
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=area3d \
  --prop view3d=20,25,15
```

**Features:** `area`, `areaStacked`, `areaPercentStacked`, `area3d`, `transparency`, `gradient`, `plotFill`, `roundedCorners`, `axisVisible`, `axisLine`

### Sheet: 5-Styling

Demonstrates styling and formatting properties on various charts.

```bash
# Fully styled chart: title effects, legend, axis fonts, series effects
officecli add data.xlsx /Sheet --type chart \
  --prop title.font=Georgia --prop title.size=18 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop title.shadow=000000-3-315-2-30 \
  --prop legendfont=10:444444:Helvetica --prop legend=right \
  --prop axisfont=9:58626E:Arial \
  --prop series.outline=FFFFFF-0.5 \
  --prop series.shadow=000000-3-315-2-25 \
  --prop roundedCorners=true --prop referenceLine=160:FF0000:1:dash

# Dual Y-axis (secondary axis)
officecli add data.xlsx /Sheet --type chart \
  --prop secondaryAxis=2

# Per-point coloring and negative value inversion
officecli add data.xlsx /Sheet --type chart \
  --prop point1.color=70AD47 --prop point3.color=FF0000 \
  --prop invertIfNeg=true

# Gradient plot fill and custom data label text
officecli add data.xlsx /Sheet --type chart \
  --prop plotFill=E8F0FE-FFFFFF:90 \
  --prop marker=diamond:8:4472C4 \
  --prop dataLabels.numFmt=#,##0 \
  --prop dataLabel3.text=Peak!
```

**Features:** `title.shadow`, `secondaryAxis`, `point{N}.color`, `invertIfNeg`, `plotFill` gradient, `dataLabels.numFmt`, `dataLabel{N}.text`

### Sheet: 6-Layout

Manual positioning and axis control properties.

```bash
# Manual layout of plot area, title, legend
officecli add data.xlsx /Sheet --type chart \
  --prop plotArea.x=0.15 --prop plotArea.y=0.15 \
  --prop plotArea.w=0.7 --prop plotArea.h=0.7 \
  --prop title.x=0.3 --prop title.y=0.01 \
  --prop legend.x=0.02 --prop legend.y=0.4 \
  --prop legend.overlay=true

# Logarithmic scale, reversed axis, display units
officecli add data.xlsx /Sheet --type chart \
  --prop logBase=10 \
  --prop axisOrientation=maxMin \
  --prop dispUnits=thousands

# Label font, separator, per-label hide
officecli add data.xlsx /Sheet --type chart \
  --prop labelFont=11:2E75B6:true \
  --prop "dataLabels.separator=: " \
  --prop dataLabel2.text=Best! \
  --prop dataLabel3.delete=true

# Error bars, minor ticks, opacity
officecli add data.xlsx /Sheet --type chart \
  --prop errBars=percentage \
  --prop majorTickMark=outside --prop minorTickMark=inside \
  --prop opacity=80
```

**Features:** `plotArea.x/y/w/h`, `title.x/y`, `legend.x/y`, `legend.overlay`, `logBase`, `axisOrientation`, `dispUnits`, `labelFont`, `dataLabels.separator`, `dataLabel{N}.delete`, `errBars`, `minorTickMark`, `opacity`

### Sheet: 7-Effects

Visual effects: gradients, conditional colors, glow, presets.

```bash
# Per-series gradients
officecli add data.xlsx /Sheet --type chart \
  --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90'

# Area fill gradient and title glow
officecli add data.xlsx /Sheet --type chart \
  --prop areafill=4472C4-BDD7EE:90 \
  --prop title.glow=4472C4-8-60

# Conditional coloring (below/above threshold)
officecli add data.xlsx /Sheet --type chart \
  --prop colorRule=60:FF0000:70AD47

# Preset style and leader lines
officecli add data.xlsx /Sheet --type chart \
  --prop style=26 \
  --prop dataLabels.showLeaderLines=true
```

**Features:** `gradients`, `areafill`, `title.glow`, `colorRule`, `style`, `dataLabels.showLeaderLines`

## Inspect the Generated File

```bash
officecli query charts-basic.xlsx chart
officecli get charts-basic.xlsx "/1-Column Charts/chart[1]"
```
