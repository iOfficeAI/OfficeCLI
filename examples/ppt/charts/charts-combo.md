# Combo Charts Showcase — combotypes, combosplit, secondaryaxis

Three files work together:

- **charts-combo.py** — build script (`python3 charts-combo.py`).
- **charts-combo.pptx** — generated deck.
- **charts-combo.md** — this file.

## Regenerate

```bash
cd examples/ppt
python3 charts-combo.py
# → charts-combo.pptx
```

## Slide map

```
  Slide 1  combotypes mixes       column+line, column+area, line+area, bar+line
  Slide 2  combosplit             split index 1, 2, 3 (first N series use primary)
  Slide 3  secondaryaxis          1 series, 2 series, multiple series on secondary
  Slide 4  Title & legend
  Slide 5  Data labels
  Slide 6  Axes                   min/max on both axes, titles, gridlines
  Slide 7  Series styling         colors, gradients, transparency, outline, shadow
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
  --prop chartType=combo \
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
- [charts-doughnut.md](charts-doughnut.md), [charts-area.md](charts-area.md), [charts-scatter.md](charts-scatter.md), [charts-bubble.md](charts-bubble.md)
- [charts-radar.md](charts-radar.md), [charts-stock.md](charts-stock.md), [charts-waterfall.md](charts-waterfall.md), [charts-3d.md](charts-3d.md)
- [excel charts examples](../excel/)
