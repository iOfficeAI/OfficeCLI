# Extended Chart Types Showcase

This demo consists of three files that work together:

- **charts-extended.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-extended.xlsx** — The generated workbook with 4 sheets (1 default + 3 chart sheets, 12 charts total).
- **charts-extended.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-extended.py
# → charts-extended.xlsx
```

## Chart Sheets

### Sheet: 1-Waterfall & Funnel

Two waterfall charts and two funnel charts covering financial bridges and sales pipelines.

```bash
# Waterfall with increase/decrease/total colors and data labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop data="Start:1000,Revenue:500,Costs:-300,Tax:-100,Net:1100" \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=FF0000 \
  --prop totalColor=4472C4 \
  --prop dataLabels=true

# Waterfall with alternative color palette and legend
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=waterfall \
  --prop data="Budget:5000,Sales:2000,Marketing:-800,Ops:-600,Net:5600" \
  --prop increaseColor=2E75B6 \
  --prop decreaseColor=C00000 \
  --prop totalColor=FFC000 \
  --prop legend=bottom

# Funnel chart — sales pipeline with descending values
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=funnel \
  --prop series1="Pipeline:1200,850,600,300,120" \
  --prop categories=Leads,Qualified,Proposal,Negotiation,Won \
  --prop dataLabels=true

# Funnel chart — marketing conversion funnel (6 stages)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=funnel \
  --prop series1="Users:10000,6500,3200,1800,900,450" \
  --prop categories=Impressions,Clicks,Signups,Active,Paying,Retained \
  --prop dataLabels=true
```

**Features:** `chartType=waterfall`, `data=` name:value pairs, `increaseColor`, `decreaseColor`, `totalColor`, `dataLabels`, `legend=bottom`, `chartType=funnel`, descending values (pipeline stages)

### Sheet: 2-Treemap & Sunburst

Two treemap charts and two sunburst charts covering hierarchical data visualization.

```bash
# Treemap with overlapping parent labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=treemap \
  --prop series1="Revenue:450,380,310,280,210,180,150,120" \
  --prop categories=Laptops,Phones,Tablets,TVs,Cameras,Audio,Gaming,Wearables \
  --prop parentLabelLayout=overlapping

# Treemap with banner-style parent labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=treemap \
  --prop series1="Budget:900,750,600,500,420,350,280" \
  --prop categories=Engineering,Sales,Marketing,Support,Finance,HR,Legal \
  --prop parentLabelLayout=banner

# Sunburst — radial hierarchical market breakdown
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=sunburst \
  --prop series1="Share:35,25,20,15,30,25,20,10,15" \
  --prop categories=North,South,East,West,Urban,Suburban,Rural,Online,Retail

# Sunburst — product category breakdown (10 segments)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=sunburst \
  --prop series1="Units:500,400,300,250,200,180,160,140,120,100" \
  --prop categories=A,B,C,D,E,F,G,H,I,J
```

**Features:** `chartType=treemap`, `parentLabelLayout=overlapping`, `parentLabelLayout=banner`, area-sized blocks, `chartType=sunburst`, radial hierarchical layout

### Sheet: 3-Histogram & BoxWhisker

Two histogram charts and two box-and-whisker charts covering statistical distributions.

```bash
# Histogram — test score distribution with auto-binning
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Scores:45,52,58,61,63,65,...,95,97,99"

# Histogram — sales distribution with explicit bin count
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Sales:120,135,...,620,700" \
  --prop binCount=5

# Box & Whisker — two-team comparison with exclusive quartiles
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=boxWhisker \
  --prop series1="TeamA:42,55,61,...,105,120" \
  --prop series2="TeamB:30,38,45,...,92,110" \
  --prop quartileMethod=exclusive

# Box & Whisker — three-department salary comparison with inclusive quartiles
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=boxWhisker \
  --prop series1="Engineering:85,92,...,150,180" \
  --prop series2="Marketing:60,65,...,98,110" \
  --prop series3="Sales:55,62,...,160,190" \
  --prop quartileMethod=inclusive
```

**Features:** `chartType=histogram`, auto-binning, `binCount`, raw value distribution, `chartType=boxWhisker`, `quartileMethod=exclusive`, `quartileMethod=inclusive`, multi-series comparison, outlier detection

## Property Coverage

| Property | Sheet |
|---|---|
| `chartType=waterfall` | 1 |
| `data=` (name:value pairs) | 1 |
| `increaseColor` | 1 |
| `decreaseColor` | 1 |
| `totalColor` | 1 |
| `dataLabels` | 1 |
| `legend` | 1 |
| `chartType=funnel` | 1 |
| `chartType=treemap` | 2 |
| `parentLabelLayout=overlapping` | 2 |
| `parentLabelLayout=banner` | 2 |
| `chartType=sunburst` | 2 |
| `chartType=histogram` | 3 |
| `binCount` | 3 |
| `chartType=boxWhisker` | 3 |
| `quartileMethod=exclusive` | 3 |
| `quartileMethod=inclusive` | 3 |

## Inspect the Generated File

```bash
officecli query charts-extended.xlsx chart
officecli get charts-extended.xlsx "/1-Waterfall & Funnel/chart[1]"
```
