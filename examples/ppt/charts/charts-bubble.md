# Bubble Charts Showcase

Three files work together:

- **charts-bubble.py** — build script (`python3 charts-bubble.py`).
- **charts-bubble.pptx** — generated deck.
- **charts-bubble.md** — this file.

## Regenerate

```bash
cd examples/ppt
python3 charts-bubble.py
# → charts-bubble.pptx
```

## Slide map

```
  Slide 1  bubbleScale            50 / 100 / 150 / 200 (% of default)
  Slide 2  sizerepresents         area vs width
  Slide 3  shownegbubbles         true vs false (with negative values)
  Slide 4  Title & legend         title.* + legend positions + legendFont
  Slide 5  Data labels            value/category/bubbleSize, labelfont
  Slide 6  Axes                   min/max, gridlines, ticks
  Slide 7  Series styling         colors, gradient, transparency, outline, shadow
  Slide 8  Presets & per-series   preset bundles + chart-series Set
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
  --prop chartType=bubble \
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

- [charts-column.md](charts-column.md), [charts-bar.md](charts-bar.md), [charts-line.md](charts-line.md), [charts-pie.md](charts-pie.md)
- [charts-doughnut.md](charts-doughnut.md), [charts-area.md](charts-area.md), [charts-scatter.md](charts-scatter.md)
- [charts-radar.md](charts-radar.md), [charts-stock.md](charts-stock.md), [charts-combo.md](charts-combo.md), [charts-waterfall.md](charts-waterfall.md), [charts-3d.md](charts-3d.md)
- [excel charts examples](../excel/)
