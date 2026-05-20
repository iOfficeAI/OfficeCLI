# Box-Whisker Chart Showcase

This demo consists of three files that work together:

- **charts-boxwhisker.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-boxwhisker.xlsx** — The generated workbook with 2 sheets (8 box-whisker charts total).
- **charts-boxwhisker.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-boxwhisker.py
# → charts-boxwhisker.xlsx
```

## Chart Sheets

### Sheet: 1-Basics & Quartile

Four box-whisker charts covering basic usage, quartile methods, title styling, and series colors.

```bash
# Chart 1: Single series, exclusive quartile, data labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=boxWhisker \
  --prop title="Test Score Distribution" \
  --prop series1="Scores:45,52,58,61,63,65,67,68,70,72,75,78,82,85,90,95,99" \
  --prop quartileMethod=exclusive \
  --prop dataLabels=true

# Chart 2: Three-series comparison, inclusive quartile, legend
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=boxWhisker \
  --prop title="Salary by Department ($k)" \
  --prop series1="Engineering:85,92,95,98,102,105,108,112,118,125,135,150,180" \
  --prop series2="Marketing:60,65,68,72,75,78,80,83,88,92,98,110" \
  --prop series3="Sales:55,62,68,75,82,90,98,105,115,125,140,160,190" \
  --prop quartileMethod=inclusive \
  --prop legend=bottom

# Chart 3: Title styling — color, size, bold, font, shadow
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=boxWhisker \
  --prop title="Styled Title Demo" \
  --prop title.color=1B2838 --prop title.size=20 \
  --prop title.bold=true --prop title.font=Georgia \
  --prop title.shadow=000000-6-45-3-50 \
  --prop series1="Data:18,22,25,28,30,32,35,38,40,42,45,55,62,78"

# Chart 4: Per-series colors and drop shadow
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=boxWhisker \
  --prop title="Custom Series Colors" \
  --prop series1="GroupA:30,38,45,52,58,62,65,68,71,74,78,85,92" \
  --prop series2="GroupB:20,28,35,40,48,55,60,66,70,80,88,95,110" \
  --prop colors=5B9BD5,ED7D31 \
  --prop series.shadow=000000-6-45-3-35
```

**Features:** `quartileMethod=exclusive`, `quartileMethod=inclusive`, `dataLabels`, `legend=bottom`, multi-series (3), `title.color`, `title.size`, `title.bold`, `title.font`, `title.shadow`, `colors` (per-series), `series.shadow`

### Sheet: 2-Axes & Styling

Four box-whisker charts covering axis control, gridlines, area fills, and a full presentation-grade chart.

```bash
# Chart 5: Axis scaling, axis titles, axis title styling, axis font
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=boxWhisker \
  --prop title="Response Time (ms)" \
  --prop series1="API:12,18,22,25,28,30,32,35,38,40,42,45,55,62,78,95,120" \
  --prop series2="DB:5,8,10,12,14,16,18,20,22,25,28,32,38,45,60" \
  --prop axismin=0 --prop axismax=130 \
  --prop majorunit=10 --prop minorunit=5 \
  --prop xAxisTitle="Service" --prop yAxisTitle="Latency (ms)" \
  --prop axisTitle.color=4A5568 --prop axisTitle.size=12 \
  --prop axisTitle.bold=true --prop axisTitle.font="Helvetica Neue" \
  --prop "axisfont=10:6B7280:Consolas"

# Chart 6: Axis visibility, axis lines, gridlines, cross-axis gridlines
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=boxWhisker \
  --prop title="Axis & Gridline Control" \
  --prop series1="Temp:15,18,20,22,24,26,28,30,32,35,38,40,42" \
  --prop cataxis.visible=false \
  --prop "valaxis.line=334155:1.5" \
  --prop gridlines=true --prop gridlineColor=E2E8F0 \
  --prop xGridlines=true --prop xGridlineColor=F1F5F9

# Chart 7: Card style — area fills/borders, gapWidth, no tick labels
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=boxWhisker \
  --prop title="Card Style" \
  --prop series1="Weight:50,55,58,60,62,64,66,68,70,72,75,78,82,88,95" \
  --prop fill=6366F1 \
  --prop gapWidth=200 \
  --prop tickLabels=false --prop gridlines=false \
  --prop plotareafill=F8FAFC --prop "plotarea.border=E2E8F0:1" \
  --prop chartareafill=FFFFFF --prop "chartarea.border=CBD5E1:0.75"

# Chart 8: Full presentation-grade — all properties combined
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=boxWhisker \
  --prop title="Server Latency Dashboard" \
  --prop title.color=0F172A --prop title.size=18 \
  --prop title.bold=true --prop title.font="Helvetica Neue" \
  --prop title.shadow=000000-4-45-2-40 \
  --prop series1="US-East:8,12,15,18,20,22,24,26,28,30,35,42,55,70,95" \
  --prop series2="EU-West:10,14,18,22,25,28,30,33,36,40,45,50,60,80" \
  --prop series3="AP-South:15,20,25,30,35,38,42,45,48,52,58,65,75,90,120" \
  --prop quartileMethod=exclusive \
  --prop colors=3B82F6,10B981,F59E0B \
  --prop series.shadow=000000-4-45-2-30 \
  --prop axismin=0 --prop axismax=130 --prop majorunit=10 \
  --prop xAxisTitle="Region" --prop yAxisTitle="Latency (ms)" \
  --prop axisTitle.color=475569 --prop axisTitle.size=11 \
  --prop axisTitle.bold=true --prop axisTitle.font="Helvetica Neue" \
  --prop "axisfont=9:64748B:Helvetica Neue" \
  --prop "axisline=CBD5E1:1" \
  --prop gridlineColor=E2E8F0 \
  --prop dataLabels=true --prop "datalabels.numfmt=0" \
  --prop legend=top --prop legend.overlay=false \
  --prop "legendfont=10:475569:Helvetica Neue" \
  --prop plotareafill=F8FAFC --prop "plotarea.border=E2E8F0:0.75" \
  --prop chartareafill=FFFFFF --prop "chartarea.border=CBD5E1:0.75"
```

**Features:** `axismin`, `axismax`, `majorunit`, `minorunit`, `xAxisTitle`, `yAxisTitle`, `axisTitle.color`, `axisTitle.size`, `axisTitle.bold`, `axisTitle.font`, `axisfont`, `cataxis.visible`, `valaxis.line`, `gridlines`, `gridlineColor`, `xGridlines`, `xGridlineColor`, `fill` (single color), `gapWidth`, `tickLabels`, `plotareafill`, `plotarea.border`, `chartareafill`, `chartarea.border`, `axisline`, `datalabels.numfmt`, `legend.overlay`, `legendfont`

## Property Coverage

| Property | Chart |
|---|---|
| `chartType=boxWhisker` | 1-8 |
| `quartileMethod=exclusive` | 1, 8 |
| `quartileMethod=inclusive` | 2 |
| `dataLabels` | 1, 8 |
| `datalabels.numfmt` | 8 |
| `legend=bottom` | 2 |
| `legend=top` | 8 |
| `legend.overlay` | 8 |
| `legendfont` | 8 |
| `title.color` | 3, 8 |
| `title.size` | 3, 8 |
| `title.bold` | 3, 8 |
| `title.font` | 3, 8 |
| `title.shadow` | 3, 8 |
| `fill` (single color) | 7 |
| `colors` (per-series) | 4, 8 |
| `series.shadow` | 4, 8 |
| `axismin` / `axismax` | 5, 8 |
| `majorunit` | 5, 8 |
| `minorunit` | 5 |
| `xAxisTitle` | 5, 8 |
| `yAxisTitle` | 5, 8 |
| `axisTitle.color` | 5, 8 |
| `axisTitle.size` | 5, 8 |
| `axisTitle.bold` | 5, 8 |
| `axisTitle.font` | 5, 8 |
| `axisfont` | 5, 8 |
| `cataxis.visible` | 6 |
| `valaxis.line` | 6 |
| `axisline` | 8 |
| `gridlines` | 6, 7 |
| `gridlineColor` | 6, 8 |
| `xGridlines` | 6 |
| `xGridlineColor` | 6 |
| `tickLabels` | 7 |
| `gapWidth` | 7 |
| `plotareafill` | 7, 8 |
| `plotarea.border` | 7, 8 |
| `chartareafill` | 7, 8 |
| `chartarea.border` | 7, 8 |

## Inspect the Generated File

```bash
officecli query charts-boxwhisker.xlsx chart
officecli get charts-boxwhisker.xlsx "/1-Basics & Quartile/chart[1]"
```
