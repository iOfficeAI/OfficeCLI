# Stock Charts Showcase

This demo consists of three files that work together:

- **charts-stock.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-stock.xlsx** — The generated workbook with 4 sheets (1 default + 3 chart sheets, 12 charts total).
- **charts-stock.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-stock.py
# -> charts-stock.xlsx
```

## Chart Sheets

### Sheet: 1-Stock Fundamentals

Four OHLC stock charts covering basic rendering, gridlines, hi-low lines, and up-down bars.

```bash
# Basic OHLC stock chart with axis titles
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop series1="Open:142,145,148,150,147,152" \
  --prop series2="High:148,151,155,156,153,158" \
  --prop series3="Low:139,142,145,147,144,149" \
  --prop series4="Close:145,148,150,147,152,155" \
  --prop catTitle=Week --prop axisTitle=Price ($)

# Stock with gridlines and axis font
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop gridlines=D9D9D9:0.5 --prop axisfont=9:666666

# Hi-low lines connecting high to low per category
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop hiLowLines=true

# Up-down bars showing open-to-close direction
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop updownbars=100:70AD47:C00000
```

**Features:** `stock`, 4-series OHLC, `catTitle`, `axisTitle`, `gridlines`, `axisfont`, `hiLowLines`, `updownbars=gapWidth:upColor:downColor`

### Sheet: 2-Stock Styling

Four styled stock charts with title fonts, axis lines, custom ranges, and chart fills.

```bash
# Title and legend styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop legend=right --prop legendfont=10:333333:Calibri

# Axis line styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop axisLine=333333-1.5 --prop catAxisLine=333333-1.5

# Custom axis range with major unit
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop axisMin=110 --prop axisMax=150 --prop majorUnit=10

# Chart area fills and rounded corners
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop plotFill=F0F4F8 --prop chartFill=FAFAFA \
  --prop roundedCorners=true
```

**Features:** `title.font/size/color/bold`, `legend=right`, `legendfont`, `axisLine`, `catAxisLine`, `axisMin/Max`, `majorUnit`, `plotFill`, `chartFill`, `roundedCorners`

### Sheet: 3-Stock Advanced

Four advanced stock charts with data labels, reference lines, borders, and number formatting.

```bash
# Data labels on stock chart
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop dataLabels=true --prop labelPos=top \
  --prop labelFont=8:666666:false

# Reference line as support/resistance level
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop referenceLine=115:Resistance:C00000

# Chart and plot area borders
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop chartArea.border=333333-1.5 \
  --prop plotArea.border=999999-0.75

# Axis number format (dollar)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=stock \
  --prop axisNumFmt=$#,##0
```

**Features:** `dataLabels`, `labelPos`, `labelFont`, `referenceLine`, `chartArea.border`, `plotArea.border`, `axisNumFmt`

## Inspect the Generated File

```bash
officecli query charts-stock.xlsx chart
officecli get charts-stock.xlsx "/1-Stock Fundamentals/chart[1]"
```
