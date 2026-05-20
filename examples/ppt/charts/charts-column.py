#!/usr/bin/env python3
"""
Column Charts Showcase — column, stackedColumn, percentStackedColumn, column3d.

Generates: charts-column.pptx

Every column-applicable property officecli exposes is demonstrated at least
once across the slides:

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

Usage:
  python3 charts-column.py
"""
import subprocess, os, sys, atexit

FILE = os.path.join(os.path.dirname(__file__), "charts-column.pptx")

def cli(*args):
    r = subprocess.run(["officecli", *args], capture_output=True, text=True)
    if r.returncode != 0:
        msg = (r.stderr or r.stdout or "").strip()
        head = msg.splitlines()[0][:160] if msg else ""
        if "UNSUPPORTED" in msg:
            # Forward-compat: skip unsupported props but surface so silent gaps are visible.
            print(f"  ⚠ {' '.join(args[:3])} → {head}", file=sys.stderr)
            return
        if msg:
            print(f"  ! {' '.join(args[:3])} → {head}", file=sys.stderr)
            sys.exit(r.returncode)

def props(d):
    out = []
    for k, v in d.items():
        out += ["--prop", f"{k}={v}"]
    return out

slide = 0
def new_slide(title):
    global slide
    slide += 1
    cli("add", FILE, "/", "--type", "slide")
    cli("add", FILE, f"/slide[{slide}]", "--type", "shape",
        *props({"text": title, "size": 24, "bold": "true",
                "autoFit": "normal", "x": "0.5in", "y": "0.3in", "width": "12.3in", "height": "0.6in"}))

def add_chart(box, p):
    p = {**box, **p}
    cli("add", FILE, f"/slide[{slide}]", "--type", "chart", *props(p))

# 2x2 grid boxes (widescreen 13.33 x 7.5in)
TL = {"x": "0.3in", "y": "1.05in",  "width": "6.1in", "height": "3in"}
TR = {"x": "6.95in","y": "1.05in",  "width": "6.1in", "height": "3in"}
BL = {"x": "0.3in", "y": "4.25in",  "width": "6.1in", "height": "3in"}
BR = {"x": "6.95in","y": "4.25in",  "width": "6.1in", "height": "3in"}
CATS = "Q1,Q2,Q3,Q4"
TWO_SERIES = "East:120,135,148,162;West:95,108,115,128"
THREE_SERIES = "East:120,135,148,162;South:95,108,115,128;West:80,90,98,110"

if os.path.exists(FILE):
    os.remove(FILE)
cli("create", FILE); cli("open", FILE)
atexit.register(lambda: (cli("close", FILE), cli("validate", FILE)))

# ---------------------------------------------------------------------------
# Slide 1 — Basic variants
# ---------------------------------------------------------------------------
new_slide("Column variants — column / stackedColumn / percentStackedColumn / column3d")
add_chart(TL, {"chartType": "column", "title": "column", "legend": "bottom",
               "categories": CATS, "data": TWO_SERIES})
add_chart(TR, {"chartType": "stackedColumn", "title": "stackedColumn", "legend": "bottom",
               "categories": CATS, "data": THREE_SERIES})
add_chart(BL, {"chartType": "percentStackedColumn", "title": "percentStackedColumn",
               "legend": "bottom", "categories": CATS, "data": THREE_SERIES})
add_chart(BR, {"chartType": "column3d", "view3d": "15,20,30", "gapdepth": "150",
               "title": "column3d (view3d=15,20,30)", "legend": "bottom",
               "categories": CATS, "data": TWO_SERIES})

# ---------------------------------------------------------------------------
# Slide 2 — Title & legend
# ---------------------------------------------------------------------------
new_slide("Title & legend — title.font/size/color/bold, legend positions, legendFont")
add_chart(TL, {"chartType": "column", "title": "Styled title",
               "title.font": "Georgia", "title.size": "20", "title.color": "4472C4",
               "title.bold": "true", "legend": "bottom",
               "categories": CATS, "data": TWO_SERIES})
add_chart(TR, {"chartType": "column", "title": "legend=top + legendFont",
               "legend": "top", "legendFont": "10:333333:Calibri",
               "categories": CATS, "data": TWO_SERIES})
add_chart(BL, {"chartType": "column", "title": "legend=topRight overlay",
               "legend": "topRight", "legend.overlay": "true",
               "categories": CATS, "data": TWO_SERIES})
add_chart(BR, {"chartType": "column", "autotitledeleted": "true", "legend": "none",
               "categories": CATS, "data": TWO_SERIES})

# ---------------------------------------------------------------------------
# Slide 3 — Data labels
# ---------------------------------------------------------------------------
new_slide("Data labels — flags (value/category/percent/none), labelPos, labelfont")
add_chart(TL, {"chartType": "column", "title": "value @ outsideEnd",
               "dataLabels": "value", "labelPos": "outsideEnd",
               "labelfont": "10:333333:Calibri", "legend": "none",
               "categories": CATS, "data": "A:60,90,140,180"})
add_chart(TR, {"chartType": "column", "title": "value,category @ insideEnd",
               "dataLabels": "value,category", "labelPos": "insideEnd",
               "labelfont": "9:FFFFFF:Calibri", "legend": "none",
               "categories": CATS, "data": "A:60,90,140,180"})
add_chart(BL, {"chartType": "stackedColumn", "title": "stacked + center labels",
               "dataLabels": "value", "labelPos": "center",
               "labelfont": "9:FFFFFF:Calibri", "legend": "bottom",
               "categories": CATS, "data": THREE_SERIES})
add_chart(BR, {"chartType": "column", "title": "dataLabels=none",
               "dataLabels": "none", "legend": "none",
               "categories": CATS, "data": "A:60,90,140,180"})

# ---------------------------------------------------------------------------
# Slide 4 — Axes
# ---------------------------------------------------------------------------
new_slide("Axes — min/max, titles, fonts, gridlines, units, log, secondary")
add_chart(TL, {"chartType": "column", "title": "axis min/max + titles + numfmt",
               "legend": "none",
               "axismin": "0", "axismax": "200", "majorunit": "50", "minorunit": "10",
               "axistitle": "Revenue (USD)", "cattitle": "Quarter",
               "axisfont": "10:333333:Calibri", "axisline": "666666:1",
               "axisnumfmt": "#,##0",
               "categories": CATS, "data": "Rev:60,90,140,180"})
add_chart(TR, {"chartType": "column", "title": "gridlines + minorGridlines + ticks",
               "legend": "none",
               "gridlines": "E0E0E0:0.3", "minorGridlines": "F0F0F0:0.25",
               "majorTickMark": "out", "minorTickMark": "in", "tickLabelPos": "nextTo",
               "labelrotation": "-30",
               "categories": "January,February,March,April",
               "data": "A:60,90,140,180"})
add_chart(BL, {"chartType": "column", "title": "dispunits=thousands",
               "legend": "none", "dispunits": "thousands",
               "categories": CATS, "data": "Rev:120000,135000,148000,162000"})
add_chart(BR, {"chartType": "combo", "combotypes": "column,line", "secondaryaxis": "2",
               "title": "secondaryaxis=2 (line on right)", "legend": "bottom",
               "categories": CATS, "data": "Sales:120,135,148,162;Growth %:5,12,18,22"})

# Post-Add chart-axis Set on first chart
cli("set", FILE, f"/slide[{slide}]/chart[1]/axis[@role=value]",
    *props({"title": "Revenue (USD)", "format": "$#,##0",
            "majorGridlines": "true", "minorGridlines": "false",
            "max": "200", "min": "0", "majorUnit": "50"}))

# ---------------------------------------------------------------------------
# Slide 5 — Series styling
# ---------------------------------------------------------------------------
new_slide("Series styling — colors, gradient(s), transparency, outline, shadow, invertifneg, colorrule")
add_chart(TL, {"chartType": "column", "title": "colors + seriesoutline",
               "legend": "bottom",
               "colors": "4472C4,ED7D31,A5A5A5",
               "seriesoutline": "000000:0.5",
               "categories": CATS, "data": THREE_SERIES})
add_chart(TR, {"chartType": "column", "title": "gradient + seriesshadow",
               "legend": "bottom",
               "gradient": "FF6600-FFCC00:90",
               "seriesshadow": "000000-5-45-3-50",
               "categories": CATS, "data": "A:60,90,140,180"})
add_chart(BL, {"chartType": "column", "title": "per-series gradients + transparency=30",
               "legend": "bottom",
               "gradients": "FF0000-0000FF;00FF00-FFFF00",
               "transparency": "30",
               "categories": CATS,
               "data": "A:60,90,140,180;B:40,70,100,130"})
add_chart(BR, {"chartType": "column", "title": "invertifneg + colorrule",
               "legend": "none",
               "invertifneg": "true",
               "colorrule": "0:FF0000:00AA00",
               "categories": "Q1,Q2,Q3,Q4,Q5",
               "data": "Net:60,-30,40,-50,80"})

# Recolor series 1 of the first chart via chart-series Set
cli("set", FILE, f"/slide[{slide}]/chart[1]/series[1]", *props({"color": "2E75B6"}))

# ---------------------------------------------------------------------------
# Slide 6 — Layout & overlays
# ---------------------------------------------------------------------------
new_slide("Layout & overlays — gapwidth, overlap, referenceline, errbars, trendline, dataTable")
add_chart(TL, {"chartType": "column", "title": "gapwidth=50 + overlap=20",
               "legend": "bottom", "gapwidth": "50", "overlap": "20",
               "categories": CATS, "data": "A:60,90,140,180;B:50,75,110,150"})
add_chart(TR, {"chartType": "column", "title": "referenceline=100",
               "legend": "none", "referenceline": "100:FF0000:Target",
               "categories": CATS, "data": "A:60,90,140,180"})
add_chart(BL, {"chartType": "column", "title": "errbars=percentage:10",
               "legend": "none", "errbars": "percentage:10",
               "categories": CATS, "data": "A:60,90,140,180"})
add_chart(BR, {"chartType": "column", "title": "dataTable=true + trendline=linear",
               "legend": "bottom", "dataTable": "true", "trendline": "linear",
               "categories": CATS, "data": "A:60,90,140,180"})

# ---------------------------------------------------------------------------
# Slide 7 — Backgrounds
# ---------------------------------------------------------------------------
new_slide("Backgrounds — chartareafill, plotFill, borders, roundedcorners")
add_chart(TL, {"chartType": "column", "title": "chartareafill + plotFill + borders",
               "legend": "bottom",
               "chartareafill": "FFF8E7", "plotFill": "FAFAFA",
               "chartborder": "000000:1", "plotborder": "CCCCCC:0.5",
               "categories": CATS, "data": "A:60,90,140,180"})
add_chart(TR, {"chartType": "column", "title": "roundedcorners=true",
               "legend": "bottom",
               "roundedcorners": "true", "chartborder": "4472C4:2",
               "categories": CATS, "data": "A:60,90,140,180"})
add_chart(BL, {"chartType": "column", "title": "plotFill=none, gridlines=none",
               "legend": "none",
               "plotFill": "none", "gridlines": "none",
               "categories": CATS, "data": "A:60,90,140,180"})
add_chart(BR, {"chartType": "column", "title": "varyColors=true (single series)",
               "legend": "none", "varyColors": "true",
               "categories": CATS, "data": "A:60,90,140,180"})

# ---------------------------------------------------------------------------
# Slide 8 — Presets & per-series control
# ---------------------------------------------------------------------------
new_slide("Presets & per-series — preset bundles + seriesN.* + chart-series Set")
add_chart(TL, {"chartType": "column", "preset": "minimal", "title": "preset=minimal",
               "legend": "bottom", "categories": CATS,
               "data": "A:60,90,140,180;B:50,75,110,150"})
add_chart(TR, {"chartType": "column", "preset": "corporate", "title": "preset=corporate",
               "legend": "bottom", "categories": CATS,
               "data": "A:60,90,140,180;B:50,75,110,150"})
add_chart(BL, {"chartType": "column", "preset": "dark", "title": "preset=dark",
               "legend": "bottom", "categories": CATS,
               "data": "A:60,90,140,180;B:50,75,110,150"})
add_chart(BR, {"chartType": "column", "title": "seriesN.* Add + chart-series Set",
               "legend": "bottom", "categories": CATS,
               "series1.name": "Product A", "series1.values": "60,90,140,180",
               "series1.color": "4472C4",
               "series2.name": "Product B", "series2.values": "50,75,110,150",
               "series2.color": "ED7D31",
               "series3.name": "Product C", "series3.values": "40,65,90,120",
               "series3.color": "70AD47"})
cli("set", FILE, f"/slide[{slide}]/chart[4]/series[1]",
    *props({"name": "Renamed Alpha", "color": "C00000"}))

print(f"Done: {FILE}  ({slide} slides)")
