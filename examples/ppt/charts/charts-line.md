# Line Charts Showcase — line, stackedLine, percentStackedLine, line3d

Three files work together:

- **charts-line.py** — build script (`python3 charts-line.py`).
- **charts-line.pptx** — generated deck.
- **charts-line.md** — this file.

## Regenerate

```bash
cd examples/ppt
python3 charts-line.py
# → charts-line.pptx
```

## Slide map

```
  Slide 1  Variants           line / stackedLine / percentStackedLine / line3d
  Slide 2  Markers            marker symbol/size/color, markersize, showMarker
  Slide 3  Smoothing & dash   smooth, linedash, linewidth
  Slide 4  Title & legend     title.* + legend positions + legendFont
  Slide 5  Data labels        flags, labelPos, labelfont
  Slide 6  Axes               min/max, titles, fonts, gridlines, ticks, labelrotation, log
  Slide 7  Overlays           droplines, hilowlines, updownbars, trendline, errbars, referenceline
  Slide 8  Per-series Set     lineWidth/lineDash/marker/markerSize/color/smooth + presets
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
  --prop chartType=line \
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

- [charts-column.md](charts-column.md), [charts-bar.md](charts-bar.md), [charts-pie.md](charts-pie.md)
- [charts-doughnut.md](charts-doughnut.md), [charts-area.md](charts-area.md), [charts-scatter.md](charts-scatter.md), [charts-bubble.md](charts-bubble.md)
- [charts-radar.md](charts-radar.md), [charts-stock.md](charts-stock.md), [charts-combo.md](charts-combo.md), [charts-waterfall.md](charts-waterfall.md), [charts-3d.md](charts-3d.md)
- [excel charts examples](../excel/)
