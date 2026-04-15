# Advanced Charts Showcase

This demo consists of three files that work together:

- **charts-advanced.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-advanced.xlsx** — The generated workbook with 3 sheets (12 charts total).
- **charts-advanced.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-advanced.py
# → charts-advanced.xlsx
```

## Chart Sheets

### Sheet: 1-Scatter & Bubble

Four charts covering scatter plot and bubble chart fundamentals.

```bash
# Scatter with circle markers and connecting lines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop categories=1,2,3,4,5,6 \
  --prop series1="SeriesA:10,25,15,40,30,50" \
  --prop series2="SeriesB:5,18,22,35,28,42" \
  --prop colors=4472C4,ED7D31 \
  --prop marker=circle --prop markerSize=8 \
  --prop lineWidth=1.5 --prop legend=bottom

# Scatter with smooth curve and reference line
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop smooth=true --prop marker=diamond --prop markerSize=7 \
  --prop referenceLine=25:FF0000:Target:dash \
  --prop axisTitle=Value --prop catTitle=Period

# Scatter with per-series marker styles
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=scatter \
  --prop series1.marker=square --prop series2.marker=triangle \
  --prop series3.marker=star --prop markerSize=9 \
  --prop lineWidth=1 --prop gridlines=D9D9D9:0.5:dot

# Bubble chart with scale control
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop bubbleScale=80 --prop legend=right \
  --prop axisTitle=Revenue --prop catTitle=Market Size
```

**Features:** `scatter`, `bubble`, `marker` (circle, diamond, square, triangle, star), `markerSize`, `series{N}.marker` (per-series), `smooth`, `lineWidth`, `referenceLine`, `bubbleScale`, `catTitle`, `axisTitle`, `gridlines`, `legend`

### Sheet: 2-Combo & Radar

Four charts covering combo (bar+line) and radar (spider) charts.

```bash
# Combo chart with comboSplit (bar+line split)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop comboSplit=2 \
  --prop series1="Revenue:120,145,132,168,155,180" \
  --prop series2="Expenses:80,92,85,98,90,105" \
  --prop series3="Growth:8,12,6,15,10,16" \
  --prop legend=bottom --prop axisTitle=Amount --prop catTitle=Month

# Combo with secondary axis
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop comboSplit=1 --prop secondaryAxis=2 \
  --prop series1="Volume:1200,1450,1320,1680" \
  --prop series2="AvgPrice:45,52,48,58"

# Combo with per-series type control (combotypes)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop combotypes=column,column,line,area

# Radar chart with radarStyle=marker
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=marker \
  --prop categories=Speed,Strength,Stamina,Agility,Accuracy \
  --prop series1="AthleteA:80,65,90,75,85" \
  --prop series2="AthleteB:70,85,60,90,70"
```

**Features:** `combo`, `comboSplit` (bar/line split point), `combotypes` (per-series type: column/line/area), `secondaryAxis`, `radar`, `radarStyle` (marker/filled/standard), `categories` as spoke labels

### Sheet: 3-Stock & Radar

Four charts covering stock (OHLC) and additional radar/bubble variants.

```bash
# Stock OHLC chart with 4 series (Open/High/Low/Close)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop categories=Mon,Tue,Wed,Thu,Fri \
  --prop series1="Open:145,148,150,147,152" \
  --prop series2="High:152,155,157,153,160" \
  --prop series3="Low:143,146,148,144,150" \
  --prop series4="Close:148,150,147,152,158" \
  --prop catTitle=Day --prop axisTitle=Price

# Stock chart — weekly OHLC with gridlines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop gridlines=E0E0E0:0.75

# Radar — filled style with transparency
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop transparency=40 --prop legend=bottom

# Bubble with single series and axis titles
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop bubbleScale=100 --prop legend=none \
  --prop axisTitle=Revenue --prop catTitle=Market Size
```

**Features:** `stock` (OHLC format: 4 series = Open/High/Low/Close), `radarStyle=filled`, `transparency` (fill alpha on radar), `bubbleScale=100`, `legend=none`, `gridlines` styling

## Complete Feature Coverage

| Feature | Sheet |
|---------|-------|
| **Chart types:** scatter, bubble, combo, radar, stock | 1, 2, 3 |
| **Scatter:** marker styles, smooth, lineWidth | 1 |
| **Bubble:** bubbleScale, single/multi-series | 1, 3 |
| **Combo:** comboSplit, combotypes, secondaryAxis | 2 |
| **Radar:** radarStyle (marker, filled, standard), transparency | 2, 3 |
| **Stock:** OHLC (4 series), gridlines | 3 |
| **Markers:** circle, diamond, square, triangle, star, per-series | 1 |
| **Data input:** inline series, categories | 1, 2, 3 |
| **Axis:** catTitle, axisTitle | 1, 2, 3 |
| **Legend:** position (bottom, right, none) | 1, 2, 3 |
| **Reference line:** value:color:label:dash | 1 |
| **Gridlines:** color:width:dash | 1, 3 |

## Inspect the Generated File

```bash
officecli query charts-advanced.xlsx chart
officecli get charts-advanced.xlsx "/1-Scatter & Bubble/chart[1]"
```
