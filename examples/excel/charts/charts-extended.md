# Extended Chart Types Showcase

This demo consists of three files that work together:

- **charts-extended.py** — Python script that calls `officecli` commands to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-extended.xlsx** — The generated workbook: 3 sheets, 14 charts, covering every property supported by the cx:chart family (waterfall, funnel, treemap, sunburst, histogram, boxWhisker).
- **charts-extended.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-extended.py
# → charts-extended.xlsx
```

## Feature Coverage Summary

Every extended-chart-specific knob is exercised by at least one chart:

| Chart type | Specific knobs | Covered by |
|---|---|---|
| waterfall | `increaseColor`, `decreaseColor`, `totalColor`, `chartFill`, `labelFont` | Sheet 1, Chart 1–2 |
| funnel | (generic styling only) | Sheet 1, Chart 3–4 |
| pareto | auto-sort desc, `ownerIdx` cumulative-% line, secondary % axis | Sheet 4, Chart 1–2 |
| treemap | `parentLabelLayout` = `overlapping` / `banner` / `none` | Sheet 2, Chart 1/2/3 |
| sunburst | (generic styling only) | Sheet 2, Chart 4 |
| histogram | `binCount`, `binSize`, `intervalClosed` = `r` / `l`, `underflowBin`, `overflowBin` | Sheet 3, Chart 1–4 |
| boxWhisker | `quartileMethod` = `exclusive` / `inclusive` | Sheet 3, Chart 5–6 |

Generic cx styling exercised across the deck: `title.glow`, `title.shadow`, `title.bold`/`size`/`color`, `dataLabels`, `labelFont`, `legend` position, `legendfont`, `axisfont`, `colors` palette, `chartFill`, `plotFill`.

> **Notes on cx:chart limitations:**
>
> - `chartFill` / `plotFill` only accept a **solid** hex color (or `none`). Unlike regular cChart, gradient `C1-C2:angle` is not supported.
> - `colors=` palette **does not work per-data-point** on single-series cx charts (funnel, treemap, sunburst). OfficeCLI only applies the first palette color to the whole series, so every bar/tile/segment ends up the same color. Omit `colors=` on these charts and let Excel's theme drive the default rainbow. `colors=` still works normally on multi-series cx charts (boxWhisker) and on all regular cChart types.

---

## Sheet: 1-Waterfall & Funnel

Two waterfall charts (financial bridges) and two funnel charts (pipelines).

```bash
# Chart 1 — waterfall with increase/decrease/total colors + data labels + title glow
officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
  --prop chartType=waterfall \
  --prop title="Cash Flow Bridge" \
  --prop data="Start:1000,Revenue:500,Costs:-300,Tax:-100,Net:1100" \
  --prop increaseColor=70AD47 --prop decreaseColor=FF0000 --prop totalColor=4472C4 \
  --prop dataLabels=true \
  --prop title.glow="00D2FF-6-60"

# Chart 2 — waterfall with legend + chartFill (solid) + custom label font
officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
  --prop chartType=waterfall \
  --prop title="Budget vs Actual" \
  --prop data="Budget:5000,Sales:2000,Marketing:-800,Ops:-600,Net:5600" \
  --prop increaseColor=2E75B6 --prop decreaseColor=C00000 --prop totalColor=FFC000 \
  --prop legend=bottom \
  --prop chartFill=F0F4FA \
  --prop dataLabels=true \
  --prop labelFont="9:333333:true"

# Chart 3 — funnel (sales pipeline) with title shadow
officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
  --prop chartType=funnel \
  --prop title="Sales Pipeline" \
  --prop series1="Pipeline:1200,850,600,300,120" \
  --prop categories=Leads,Qualified,Proposal,Negotiation,Won \
  --prop dataLabels=true \
  --prop title.shadow="000000-4-45-2-40"

# Chart 4 — funnel (marketing) with custom colors palette, legend/axis fonts
officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
  --prop chartType=funnel \
  --prop title="Marketing Funnel" \
  --prop series1="Users:10000,6500,3200,1800,900,450" \
  --prop categories=Impressions,Clicks,Signups,Active,Paying,Retained \
  --prop dataLabels=true \
  --prop legendfont="9:8B949E:Helvetica Neue" \
  --prop axisfont="10:58626E:Helvetica Neue"
```

**Features:** `chartType=waterfall`, `increaseColor`, `decreaseColor`, `totalColor`, `chartType=funnel`, descending pipeline values, `dataLabels`, `title.glow`, `title.shadow`, `legend=bottom`, `chartFill` (solid hex), `labelFont`, `colors` palette, `legendfont`, `axisfont`.

---

## Sheet: 2-Treemap & Sunburst

Three treemaps (one per `parentLabelLayout` value) and one sunburst.

```bash
# Chart 1 — treemap with parentLabelLayout=overlapping + dataLabels
officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
  --prop chartType=treemap \
  --prop title="Revenue by Product" \
  --prop series1="Revenue:450,380,310,280,210,180,150,120" \
  --prop categories=Laptops,Phones,Tablets,TVs,Cameras,Audio,Gaming,Wearables \
  --prop parentLabelLayout=overlapping \
  --prop dataLabels=true

# Chart 2 — treemap with parentLabelLayout=banner + title styling
officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
  --prop chartType=treemap \
  --prop title="Department Budget" \
  --prop series1="Budget:900,750,600,500,420,350,280" \
  --prop categories=Engineering,Sales,Marketing,Support,Finance,HR,Legal \
  --prop parentLabelLayout=banner \
  --prop title.bold=true --prop title.size=14 --prop title.color=2E5090

# Chart 3 — treemap with parentLabelLayout=none (flat, no parent header strip)
officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
  --prop chartType=treemap \
  --prop title="Flat Treemap (no parent labels)" \
  --prop series1="Units:250,200,180,160,140,120,100,80,60,40" \
  --prop categories=A,B,C,D,E,F,G,H,I,J \
  --prop parentLabelLayout=none \
  --prop dataLabels=true

# Chart 4 — sunburst with chartFill + plotFill (solid) + colors palette
officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
  --prop chartType=sunburst \
  --prop title="Market Share by Region" \
  --prop series1="Share:35,25,20,15,30,25,20,10,15" \
  --prop categories=North,South,East,West,Urban,Suburban,Rural,Online,Retail \
  --prop chartFill=F8FAFC --prop plotFill=FFFFFF \
  --prop dataLabels=true
```

**Features:** `chartType=treemap`, `parentLabelLayout=overlapping`, `parentLabelLayout=banner`, `parentLabelLayout=none`, `chartType=sunburst`, radial hierarchical layout, `colors` palette, `title.bold`/`size`/`color`, `dataLabels`, `chartFill` + `plotFill` (solid).

---

## Sheet: 3-Histogram & BoxWhisker

Four histograms covering every binning knob, and two box-and-whisker charts (one per quartile method).

```bash
# Chart 1 — histogram with auto-binning (no binCount/binSize)
officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=histogram \
  --prop title="Test Scores (auto bins)" \
  --prop series1="Scores:45,52,58,61,63,...,95,97,99"

# Chart 2 — histogram with explicit binCount=5 + title glow
officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=histogram \
  --prop title="Sales (binCount=5)" \
  --prop series1="Sales:120,135,...,620,700" \
  --prop binCount=5 \
  --prop title.glow="FFC000-6-50"

# Chart 3 — histogram with explicit binSize=50 (fixed bin width) + label font
officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=histogram \
  --prop title="Sales (binSize=50)" \
  --prop series1="Sales:120,135,...,620,700" \
  --prop binSize=50 \
  --prop dataLabels=true --prop labelFont="9:FFFFFF:true"

# Chart 4 — histogram with underflowBin + overflowBin + intervalClosed=l
officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=histogram \
  --prop title="Response Time (outlier bins)" \
  --prop series1="ms:40,55,68,75,...,220,280,350" \
  --prop underflowBin=60 \
  --prop overflowBin=200 \
  --prop intervalClosed=l \
  --prop dataLabels=true \
  --prop legend=none

# Chart 5 — box & whisker, two teams, quartileMethod=exclusive
officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Response Time by Team (ms)" \
  --prop series1="TeamA:42,55,...,105,120" \
  --prop series2="TeamB:30,38,...,92,110" \
  --prop quartileMethod=exclusive \
  --prop legend=bottom

# Chart 6 — box & whisker, three departments, quartileMethod=inclusive + title glow
officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Salary Distribution (\$k)" \
  --prop series1="Engineering:85,92,...,150,180" \
  --prop series2="Marketing:60,65,...,98,110" \
  --prop series3="Sales:55,62,...,160,190" \
  --prop quartileMethod=inclusive \
  --prop title.glow="00D2FF-6-60" \
  --prop legend=bottom
```

**Features:** `chartType=histogram`, auto-binning, `binCount` (explicit count), `binSize` (explicit width — mutually exclusive with `binCount`), `underflowBin` (cutoff for `<N`), `overflowBin` (cutoff for `>N`), `intervalClosed=r` (default, `(a,b]`) vs `intervalClosed=l` (`[a,b)`), `chartType=boxWhisker`, `quartileMethod=exclusive`, `quartileMethod=inclusive`, multi-series grouping (2 or 3), `title.glow`, `legend=bottom`, `legend=none`, `labelFont`, `dataLabels`.

---

## Sheet: 4-Pareto

Two Pareto charts demonstrating automatic descending sort and cumulative-% overlay line.

```bash
# Chart 1 — categorical Pareto (defect analysis), pre-sorted input
officecli add charts-extended.xlsx "/4-Pareto" --type chart \
  --prop chartType=pareto \
  --prop title="Defect Pareto" \
  --prop series1="Count:45,30,10,8,5,2" \
  --prop categories=Scratches,Dents,Cracks,Chips,Stains,Other \
  --prop dataLabels=true

# Chart 2 — Pareto with out-of-order input (auto-sorted desc by officecli)
officecli add charts-extended.xlsx "/4-Pareto" --type chart \
  --prop chartType=pareto \
  --prop title="Root Cause Pareto" \
  --prop series1="Tickets:12,87,5,45,3,120,22,67,8,31" \
  --prop categories=Network,Auth,DB,Cache,UI,Config,Deploy,Monitor,Queue,Storage \
  --prop title.glow="FFC000-6-50" \
  --prop legend=bottom
```

**Features:** `chartType=pareto`, automatic descending sort of values + categories, cumulative-% overlay line on secondary 0-100% axis (auto-generated via `ownerIdx`), `dataLabels`, `title.glow`, `legend=bottom`. Input is a SINGLE user series; officecli synthesizes the 2-series structure internally (clusteredColumn bars + paretoLine with `ownerIdx="0"` + secondary percentage axis).

---

## Property Reference

| Property | Applies to | Example value | Sheet |
|---|---|---|---|
| `chartType=waterfall` | waterfall | `waterfall` | 1 |
| `chartType=funnel` | funnel | `funnel` | 1 |
| `chartType=treemap` | treemap | `treemap` | 2 |
| `chartType=sunburst` | sunburst | `sunburst` | 2 |
| `chartType=histogram` | histogram | `histogram` | 3 |
| `chartType=boxWhisker` | boxWhisker | `boxWhisker` | 3 |
| `chartType=pareto` | pareto | `pareto` | 4 |
| `data=` name:value pairs | waterfall | `Start:1000,Revenue:500,...` | 1 |
| `increaseColor` | waterfall | `70AD47` | 1 |
| `decreaseColor` | waterfall | `FF0000` | 1 |
| `totalColor` | waterfall | `4472C4` | 1 |
| `series1=Name:values`, `series2=...`, `series3=...` | all cx | `TeamA:42,55,...` | 1/2/3 |
| `categories` | all cx except histogram | `Leads,Qualified,...` | 1/2 |
| `parentLabelLayout` | treemap | `overlapping` \| `banner` \| `none` | 2 |
| `binCount` | histogram | `5` | 3 |
| `binSize` | histogram | `50` | 3 |
| `intervalClosed` | histogram | `r` (default) \| `l` | 3 |
| `underflowBin` | histogram | `60` | 3 |
| `overflowBin` | histogram | `200` | 3 |
| `quartileMethod` | boxWhisker | `exclusive` \| `inclusive` | 3 |
| `dataLabels` | all cx | `true` | 1/2/3 |
| `labelFont` | all cx | `"9:FFFFFF:true"` | 1/3 |
| `title.glow` | all cx | `"00D2FF-6-60"` | 1/3 |
| `title.shadow` | all cx | `"000000-4-45-2-40"` | 1 |
| `title.bold`/`size`/`color` | all cx | `true` / `14` / `2E5090` | 2 |
| `legend` | all cx | `bottom` \| `none` | 1/3 |
| `legendfont` | all cx | `"9:8B949E:Helvetica Neue"` | 1 |
| `axisfont` | all cx | `"10:58626E:Helvetica Neue"` | 1 |
| `colors` | multi-series cx only (not useful on funnel/treemap/sunburst — see limitations note) | `4472C4,5B9BD5,...` | — |
| `chartFill` (solid only) | all cx | `F8FAFC` | 1/2 |
| `plotFill` (solid only) | all cx | `FFFFFF` | 2 |

---

## Known Validation Warning

`officecli validate charts-extended.xlsx` reports schema warnings on histogram charts' `binCount` / `binSize` elements:

```
[Schema] The element '...:binCount' has invalid value ''. The text value cannot be empty.
[Schema] The 'val' attribute is not declared.
```

This is expected. The Open XML SDK's generated schema models `cx:binCount` as a text-valued leaf (`<binCount>5</binCount>`), but **real Excel writes and requires** the attribute form (`<binCount val="5"/>`). OfficeCLI writes the Excel-compatible form via a raw unknown element; the SDK validator then complains. See `ChartExBuilder.cs:793–801` for the rationale. Files open and render correctly in Excel.

---

## Inspect the Generated File

```bash
officecli query charts-extended.xlsx chart
officecli get charts-extended.xlsx "/1-Waterfall & Funnel/chart[1]"
officecli view charts-extended.xlsx outline
```
