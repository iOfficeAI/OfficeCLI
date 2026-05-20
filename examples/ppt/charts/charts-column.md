# Column Charts Showcase — column, stackedColumn, percentStackedColumn, column3d

Three files work together:

- **charts-column.py** — build script (`python3 charts-column.py`).
- **charts-column.pptx** — generated deck.
- **charts-column.md** — this file.

## Regenerate

```bash
cd examples/ppt
python3 charts-column.py
# → charts-column.pptx
```

## Slide map

```
  Slide 1  Basic variants     column / stackedColumn / percentStackedColumn / column3d
  Slide 2  Title & legend     title.font/size/color/bold, legend positions, legendFont
  Slide 3  Data labels        dataLabels flags, labelPos, labelfont
  Slide 4  Axes               axismin/max, axistitle, axisfont, axisline, axisnumfmt,
                              gridlines, minorGridlines, majorunit, minorunit, labelrotation,
                              dispunits, logbase, secondaryaxis, chart-axis Set
  Slide 5  Series styling     colors, gradient, gradients, transparency, seriesoutline,
                              seriesshadow, invertifneg, colorrule
  Slide 6  Layout & overlays  gapwidth, overlap, referenceline, errbars, trendline, dataTable
  Slide 7  Backgrounds        chartareafill, plotFill, chartborder, plotborder, roundedcorners
  Slide 8  Presets & per-ser  preset bundles + seriesN.name/values/color + chart-series Set
```

## Reference

```bash
officecli help pptx chart            # chart-level properties
officecli help pptx chart-series     # per-series Set/Get
officecli help pptx chart-axis       # per-axis Set/Get (after creation)
```

## Pattern

```bash
# Create a chart on a slide
officecli add deck.pptx /slide[1] --type chart \
  --prop chartType=column \
  --prop x=1in --prop y=1in --prop width=10in --prop height=5in \
  --prop title="Example" --prop legend=bottom \
  --prop categories="Q1,Q2,Q3,Q4" \
  --prop data="A:60,90,140,180;B:50,75,110,150"

# Mutate a series after creation
officecli set deck.pptx /slide[1]/chart[1]/series[1] \
  --prop name="Renamed" --prop color=C00000

# Mutate an axis after creation
officecli set deck.pptx /slide[1]/chart[1]/axis[@role=value] \
  --prop title="USD" --prop format="\$#,##0" \
  --prop majorGridlines=true --prop min=0 --prop max=200
```

## Related

- [charts-bar.md](charts-bar.md), [charts-line.md](charts-line.md), [charts-pie.md](charts-pie.md)
- [charts-doughnut.md](charts-doughnut.md), [charts-area.md](charts-area.md), [charts-scatter.md](charts-scatter.md), [charts-bubble.md](charts-bubble.md)
- [charts-radar.md](charts-radar.md), [charts-stock.md](charts-stock.md), [charts-combo.md](charts-combo.md), [charts-waterfall.md](charts-waterfall.md), [charts-3d.md](charts-3d.md)
- [excel charts examples](../excel/)
