# Bubble Charts Showcase

This demo consists of three files that work together:

- **charts-bubble.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-bubble.xlsx** — The generated workbook with 4 sheets (1 default + 3 chart sheets, 12 charts total).
- **charts-bubble.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-bubble.py
# -> charts-bubble.xlsx
```

## Chart Sheets

### Sheet: 1-Bubble Fundamentals

Four bubble charts covering basic rendering, bubble scale, size representation, and data labels.

```bash
# Basic bubble with 2 series (X,Y,Size triplets separated by semicolons)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop series1="Enterprise:50,12,80;120,8,45;200,15,60" \
  --prop series2="Consumer:30,25,50;80,18,35;150,22,70" \
  --prop catTitle=Market Size ($M) --prop axisTitle=Growth Rate (%)

# bubbleScale=100 with center data labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop bubbleScale=100 \
  --prop dataLabels=true --prop labelPos=center

# Small bubbles with bubbleScale=50
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop bubbleScale=50

# Size proportional to diameter (width) instead of area
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop sizeRepresents=width
```

**Features:** `bubble`, X;Y;Size triplet format, `catTitle`, `axisTitle`, `bubbleScale`, `dataLabels`, `labelPos=center`, `labelFont`, `sizeRepresents=width`

### Sheet: 2-Bubble Styling

Four styled bubble charts with title fonts, transparency, grid styling, and shadow effects.

```bash
# Title and legend styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop legend=right --prop legendfont=10:333333:Calibri

# Transparent overlapping bubbles (ARGB with alpha)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop colors=804472C4,80ED7D31 \
  --prop bubbleScale=120

# Grid and axis line styling
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop gridlines=D9D9D9:0.5 --prop axisfont=9:666666 \
  --prop axisLine=333333-1

# Shadow and fill effects
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop plotFill=F0F4F8 --prop chartFill=FAFAFA \
  --prop series.shadow=000000-4-315-2-30
```

**Features:** `title.font/size/color/bold`, `legend=right`, `legendfont`, ARGB transparency (`80RRGGBB`), `bubbleScale`, `gridlines`, `axisfont`, `axisLine`, `plotFill`, `chartFill`, `series.shadow`

### Sheet: 3-Bubble Advanced

Four advanced bubble charts with secondary axis, reference lines, log scale, and trendlines.

```bash
# Secondary axis for second series
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop secondaryAxis=2

# Reference line (growth threshold)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop referenceLine=18:Target Growth:C00000

# Logarithmic scale with axis range
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop axisMin=1 --prop axisMax=50 \
  --prop logBase=10

# Borders and trendline
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=bubble \
  --prop chartArea.border=333333-1.5 \
  --prop plotArea.border=999999-0.75 \
  --prop trendline=linear
```

**Features:** `secondaryAxis`, `referenceLine`, `axisMin/Max`, `logBase`, `chartArea.border`, `plotArea.border`, `trendline=linear`

## Inspect the Generated File

```bash
officecli query charts-bubble.xlsx chart
officecli get charts-bubble.xlsx "/1-Bubble Fundamentals/chart[1]"
```
