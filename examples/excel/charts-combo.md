# Combo Charts Showcase

This demo consists of three files that work together:

- **charts-combo.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-combo.xlsx** — The generated workbook with 5 sheets (1 default + 4 chart sheets, 16 charts total).
- **charts-combo.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-combo.py
# -> charts-combo.xlsx
```

## Chart Sheets

### Sheet: 1-Combo Fundamentals

Four combo charts covering comboSplit, secondaryAxis, combotypes, and combined usage.

```bash
# Basic combo: 2 bar series + 1 line via comboSplit
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop series1="Revenue:120,145,160,180,195" \
  --prop series2="Expenses:90,100,110,115,125" \
  --prop series3="Margin %:25,31,31,36,36" \
  --prop comboSplit=2 --prop legend=bottom

# Combo with secondary Y-axis for line series
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop comboSplit=1 --prop secondaryAxis=2 \
  --prop catTitle=Year --prop axisTitle=Sales ($K)

# Per-series type control via combotypes
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop combotypes=column,column,line,area

# combotypes + secondaryAxis together
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop combotypes=column,column,line \
  --prop secondaryAxis=3
```

**Features:** `combo`, `comboSplit`, `secondaryAxis`, `combotypes=column,column,line,area`, `catTitle`, `axisTitle`

### Sheet: 2-Combo Styling

Four styled combo charts with title fonts, gradients, data labels, and chart fills.

```bash
# Title, legend, axis font styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop legendfont=10:333333:Calibri --prop axisfont=9:666666

# Series shadow and gradients
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop 'gradients=1F4E79-5B9BD5:90;C55A11-F4B183:90' \
  --prop series.shadow=000000-4-315-2-30

# Data labels on combo series
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop dataLabels=true --prop labelPos=top \
  --prop labelFont=9:333333:true

# Chart area styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop plotFill=F0F4F8 --prop chartFill=FAFAFA \
  --prop roundedCorners=true
```

**Features:** `title.font/size/color/bold`, `legendfont`, `axisfont`, `gradients`, `series.shadow`, `dataLabels`, `labelPos`, `labelFont`, `plotFill`, `chartFill`, `roundedCorners`

### Sheet: 3-Combo Advanced

Four advanced combo charts with reference lines, axis scaling, layout, and markers.

```bash
# Reference line and gridlines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop referenceLine=110:Target:C00000 \
  --prop gridlines=D9D9D9:0.5

# Axis scaling and display units
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop axisMin=1000000 --prop axisMax=2000000 \
  --prop dispUnits=thousands

# Manual plot layout
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop plotLayout=0.1,0.15,0.85,0.75

# Multiple line series with markers
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop comboSplit=1 --prop secondaryAxis=2,3,4 \
  --prop markers=circle-6
```

**Features:** `referenceLine`, `gridlines`, `axisMin/Max`, `dispUnits`, `plotLayout`, `markers`, multiple secondary axis series

### Sheet: 4-Combo Effects

Four effect-heavy combo charts with glow, borders, color rules, and complex multi-series.

```bash
# Title glow and shadow effects
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop title.glow=4472C4-6 \
  --prop title.shadow=000000-3-315-2-30

# Chart and plot area borders
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop chartArea.border=333333-1.5 \
  --prop plotArea.border=999999-0.75

# Color rule (conditional bar coloring)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop colorRule=80:C00000:70AD47

# 5-series dashboard with mixed combotypes
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=combo \
  --prop combotypes=column,column,column,area,line \
  --prop secondaryAxis=5
```

**Features:** `title.glow`, `title.shadow`, `chartArea.border`, `plotArea.border`, `colorRule`, 5-series `combotypes`

## Inspect the Generated File

```bash
officecli query charts-combo.xlsx chart
officecli get charts-combo.xlsx "/1-Combo Fundamentals/chart[1]"
```
