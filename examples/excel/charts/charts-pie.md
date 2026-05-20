# Pie & Doughnut Charts Showcase

This demo consists of three files that work together:

- **charts-pie.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-pie.xlsx** — The generated workbook with 3 sheets (1 default + 2 chart sheets, 8 charts total).
- **charts-pie.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-pie.py
# → charts-pie.xlsx
```

## Chart Sheets

### Sheet: 1-Pie Charts

Four pie chart variants covering flat, 3D, exploded, and gradient fills.

```bash
# Basic pie with colors and data labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=pie \
  --prop series1="Share:40,25,20,15" \
  --prop categories=Product A,Product B,Product C,Product D \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop dataLabels=true --prop labelPos=outsideEnd

# Exploded pie with per-point colors and percentage labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=pie \
  --prop explosion=15 \
  --prop point1.color=1F4E79 --prop point2.color=2E75B6 \
  --prop dataLabels.numFmt=0.0"%" --prop labelPos=bestFit

# 3D pie with tilt angle and styled title
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=pie3d \
  --prop view3d=30,0,0 \
  --prop title.font=Georgia --prop title.size=16 \
  --prop labelFont=12:FFFFFF:true --prop labelPos=center

# Pie with per-slice gradients and leader lines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=pie \
  --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90;...' \
  --prop dataLabels.showLeaderLines=true \
  --prop legend=right --prop legendfont=10:333333:Helvetica
```

**Features:** `pie`, `pie3d`, `explosion`, `point{N}.color`, `view3d`, `labelPos=bestFit`, `dataLabels.numFmt`, `labelFont`, `title.font/size/color/bold`, `gradients` (per-slice), `dataLabels.showLeaderLines`, `legendfont`, `chartFill`, `roundedCorners`

### Sheet: 2-Doughnut Charts

Four doughnut chart variants including multi-ring and styled effects.

```bash
# Basic doughnut with center labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=doughnut \
  --prop dataLabels=true --prop labelPos=center \
  --prop labelFont=14:FFFFFF:true

# Multi-ring doughnut (multiple series = concentric rings)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=doughnut \
  --prop series1="2024:40,35,25" \
  --prop series2="2025:45,30,25" \
  --prop series.outline=FFFFFF-1

# Styled doughnut with shadow effects
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=doughnut \
  --prop series.shadow=000000-4-315-2-30 \
  --prop title.shadow=000000-3-315-2-30 \
  --prop plotFill=F5F5F5

# Doughnut with explosion and per-slice gradients
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=doughnut \
  --prop explosion=8 \
  --prop 'gradients=1F4E79-5B9BD5:90;C55A11-F4B183:90;...'
```

**Features:** `doughnut`, multi-ring (multiple `series`), `labelPos=center`, `labelFont`, `series.outline`, `series.shadow`, `title.shadow`, `plotFill`, `explosion`, `gradients`

## Inspect the Generated File

```bash
officecli query charts-pie.xlsx chart
officecli get charts-pie.xlsx "/1-Pie Charts/chart[1]"
```
