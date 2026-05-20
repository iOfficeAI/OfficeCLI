# 3D Charts Showcase — column3d / bar3d / pie3d / line3d / area3d with view3d, gapdepth, shape

Three files work together:

- **charts-3d.py** — build script (`python3 charts-3d.py`).
- **charts-3d.pptx** — generated deck.
- **charts-3d.md** — this file.

## Regenerate

```bash
cd examples/ppt
python3 charts-3d.py
# → charts-3d.pptx
```

## Slide map

```
  Slide 1  3D families            column3d / bar3d / pie3d / line3d
  Slide 2  area3d & stacked 3D    area3d / stackedColumn3d / percentStackedColumn3d / line3d stacked
  Slide 3  view3d                 different rotX,rotY,perspective angles
  Slide 4  gapdepth               0 / 50 / 150 / 300 (3D bar/column/line/area only)
  Slide 5  bar shapes             box / cylinder / cone / pyramid (bar3d / column3d)
  Slide 6  Title & legend
  Slide 7  Series styling         colors, gradient, transparency, outline, shadow
  Slide 8  Presets
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
  --prop chartType=column3d \
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
- [charts-radar.md](charts-radar.md), [charts-stock.md](charts-stock.md), [charts-combo.md](charts-combo.md), [charts-waterfall.md](charts-waterfall.md)
- [excel charts examples](../excel/)
