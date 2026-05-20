# Radar Charts Showcase

This demo consists of three files that work together:

- **charts-radar.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-radar.xlsx** — The generated workbook with 5 sheets (1 default + 4 chart sheets, 16 charts total).
- **charts-radar.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-radar.py
# → charts-radar.xlsx
```

## Chart Sheets

### Sheet: 1-Radar Fundamentals

Four radar chart variants covering standard, marker, and filled styles.

```bash
# Basic radar (standard) with 3 series
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=standard \
  --prop series1="Alice:85,70,90,60,75" \
  --prop series2="Bob:65,90,70,80,85" \
  --prop categories=Speed,Strength,Stamina,Agility,Accuracy \
  --prop colors=4472C4,ED7D31,70AD47

# Radar with markers and data labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=marker \
  --prop marker=circle:6:2E75B6 \
  --prop dataLabels=true

# Filled radar with transparency
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop transparency=40

# Filled radar with white outline separators
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop series.outline=FFFFFF-0.5 \
  --prop transparency=35
```

**Features:** `radar`, `radarStyle=standard/marker/filled`, `marker=circle:6:color`, `transparency`, `series.outline`, `dataLabels`, `legend=bottom`

### Sheet: 2-Radar Styling

Four charts demonstrating title styling, shadows, axis fonts, gridlines, and chart area decoration.

```bash
# Title styling with font, size, color, bold, shadow
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop title.font=Georgia --prop title.size=18 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop title.shadow=000000-3-315-2-30

# Series shadow on filled radar
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop series.shadow=000000-4-315-2-30 \
  --prop transparency=30

# Axis font and gridlines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop axisfont=10:333333:Calibri \
  --prop gridlines=D9D9D9:0.5

# Chart area styling with fills, corners, borders
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop plotFill=F5F5F5 --prop chartFill=FAFAFA \
  --prop roundedCorners=true \
  --prop chartArea.border=BFBFBF:0.5 \
  --prop plotArea.border=D9D9D9:0.25
```

**Features:** `title.font/size/color/bold/shadow`, `series.shadow`, `axisfont`, `gridlines`, `plotFill`, `chartFill`, `roundedCorners`, `chartArea.border`, `plotArea.border`

### Sheet: 3-Labels & Legend

Four charts covering data labels, legend positioning, manual layout, and multi-series comparison.

```bash
# Data labels with font styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=marker \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop labelFont=9:333333:true \
  --prop marker=circle:6:2E75B6

# Legend positioning with overlay
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop legend=right \
  --prop legendfont=10:1F4E79:Calibri \
  --prop legend.overlay=true

# Manual plot area layout (fractional)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop plotArea.x=0.1 --prop plotArea.y=0.15 \
  --prop plotArea.w=0.8 --prop plotArea.h=0.75

# Five series comparison
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop series1="Dev:90,70,80,65,75" \
  --prop series2="QA:60,85,70,80,90" \
  --prop series3="Design:75,80,85,70,60" \
  --prop series4="PM:80,65,75,90,70" \
  --prop series5="DevOps:70,75,60,85,80" \
  --prop colors=4472C4,ED7D31,70AD47,FFC000,7030A0
```

**Features:** `dataLabels`, `labelPos=outsideEnd`, `labelFont`, `legend=right`, `legendfont`, `legend.overlay`, `plotArea.x/y/w/h`, 5+ series on single radar

### Sheet: 4-Advanced

Four charts with advanced effects: title glow, many-spoke layouts, themed styling, and overlap visualization.

```bash
# Title with glow and shadow effects
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop title.glow=4472C4-8 \
  --prop title.shadow=000000-3-315-2-30 \
  --prop marker=diamond:7:2E75B6

# 8-spoke radar with benchmark overlay
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop categories=Technical,Communication,Leadership,Creativity,Analytical,Teamwork,Adaptability,Initiative \
  --prop gridlines=D9D9D9:0.25 --prop transparency=35

# Single-series with themed purple styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=marker \
  --prop colors=7030A0 --prop marker=square:7:7030A0 \
  --prop title.color=7030A0 --prop plotFill=F8F0FF \
  --prop chartArea.border=7030A0:0.5 --prop roundedCorners=true

# Before/After comparison with low transparency overlap
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop transparency=20 \
  --prop series.outline=FFFFFF-0.75 \
  --prop chartFill=FAFAFA --prop plotFill=F5F5F5
```

**Features:** `title.glow`, `title.shadow`, `marker=diamond/square`, 8-category spokes, themed color scheme, low-transparency overlap visualization, before/after comparison pattern

## Inspect the Generated File

```bash
officecli query charts-radar.xlsx chart
officecli get charts-radar.xlsx "/1-Radar Fundamentals/chart[1]"
```
