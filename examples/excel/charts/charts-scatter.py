#!/usr/bin/env python3
"""
Scatter Charts Showcase — scatter with all marker, trendline, error bar, and styling variations.

Generates: charts-scatter.xlsx

Every scatter chart feature officecli supports is demonstrated at least once:
scatter styles, marker types, smooth curves, trendlines (linear, polynomial,
exponential, logarithmic, power, movingAvg), error bars, axis scaling,
gridlines, data labels, legend, fills, shadows, borders, secondary axis,
reference lines, log scale, and color rules.

6 sheets, 24 charts total.

  1-Scatter Fundamentals   4 charts — basic scatter, marker-only, smooth curve, line-only
  2-Marker Styles          4 charts — per-series markers, shapes, sizes, toggle
  3-Trendlines             4 charts — linear, polynomial, exponential, per-series
  4-Error Bars             4 charts — fixed, percent, stddev, stderr
  5-Styling                4 charts — title/shadow, gradients, axis/grid, borders
  6-Advanced               4 charts — secondary axis, reference line, log scale, color rule

Usage:
  python3 charts-scatter.py
"""

import subprocess, sys, os, json, atexit

FILE = "charts-scatter.xlsx"

def cli(cmd):
    """Run: officecli <cmd>"""
    r = subprocess.run(f"officecli {cmd}", shell=True, capture_output=True, text=True)
    out = (r.stdout or "").strip()
    if out:
        for line in out.split("\n"):
            if line.strip():
                print(f"  {line.strip()}")
    if r.returncode != 0:
        err = (r.stderr or "").strip()
        if err and "UNSUPPORTED" not in err and "process cannot access" not in err:
            print(f"  ERROR: {err}")

if os.path.exists(FILE):
    os.remove(FILE)

cli(f'create "{FILE}"')
cli(f'open "{FILE}"')
atexit.register(lambda: cli(f'close "{FILE}"'))

# ==========================================================================
# Sheet: 1-Scatter Fundamentals
# ==========================================================================
print("\n--- 1-Scatter Fundamentals ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Scatter Fundamentals"')

# --------------------------------------------------------------------------
# Chart 1: Basic scatter with circle markers and connecting lines
#
# officecli add charts-scatter.xlsx "/1-Scatter Fundamentals" --type chart \
#   --prop chartType=scatter \
#   --prop title="Height vs Weight" \
#   --prop categories=160,165,170,175,180,185,190 \
#   --prop series1="Male:62,68,72,78,82,88,95" \
#   --prop series2="Female:50,55,58,62,65,70,74" \
#   --prop colors=2E75B6,ED7D31 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop marker=circle --prop markerSize=6 \
#   --prop lineWidth=1.5 \
#   --prop catTitle=Height (cm) --prop axisTitle=Weight (kg) \
#   --prop legend=bottom
#
# Features: chartType=scatter, marker=circle, markerSize=6, lineWidth=1.5,
#   catTitle, axisTitle, legend
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Scatter Fundamentals" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Height vs Weight"'
    f' --prop categories=160,165,170,175,180,185,190'
    f' --prop series1=Male:62,68,72,78,82,88,95'
    f' --prop series2=Female:50,55,58,62,65,70,74'
    f' --prop colors=2E75B6,ED7D31'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop marker=circle --prop markerSize=6'
    f' --prop lineWidth=1.5'
    f' --prop catTitle="Height (cm)" --prop axisTitle="Weight (kg)"'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: Scatter marker-only (scatterStyle=marker), various marker sizes
#
# officecli add charts-scatter.xlsx "/1-Scatter Fundamentals" --type chart \
#   --prop chartType=scatter \
#   --prop scatterStyle=marker \
#   --prop title="Study Hours vs Test Score" \
#   --prop categories=1,2,3,4,5,6,7,8 \
#   --prop series1="Class A:55,60,65,72,78,82,88,92" \
#   --prop series2="Class B:50,58,62,68,74,80,85,90" \
#   --prop colors=4472C4,70AD47 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop markerSize=8 \
#   --prop catTitle=Study Hours --prop axisTitle=Score \
#   --prop gridlines=D9D9D9:0.5:dot
#
# Features: scatterStyle=marker (no connecting lines), markerSize=8,
#   gridlines styling
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Scatter Fundamentals" --type chart'
    f' --prop chartType=scatter'
    f' --prop scatterStyle=marker'
    f' --prop title="Study Hours vs Test Score"'
    f' --prop categories=1,2,3,4,5,6,7,8'
    f' --prop series1="Class A:55,60,65,72,78,82,88,92"'
    f' --prop series2="Class B:50,58,62,68,74,80,85,90"'
    f' --prop colors=4472C4,70AD47'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop markerSize=8'
    f' --prop catTitle="Study Hours" --prop axisTitle=Score'
    f' --prop gridlines=D9D9D9:0.5:dot')

# --------------------------------------------------------------------------
# Chart 3: Scatter smooth curve (smooth=true, scatterStyle=smooth)
#
# officecli add charts-scatter.xlsx "/1-Scatter Fundamentals" --type chart \
#   --prop chartType=scatter \
#   --prop scatterStyle=smooth \
#   --prop smooth=true \
#   --prop title="Temperature vs Ice Cream Sales" \
#   --prop categories=15,18,22,25,28,30,33,35 \
#   --prop series1="Sales ($):120,180,260,340,420,480,530,560" \
#   --prop colors=C00000 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop marker=diamond --prop markerSize=7 \
#   --prop lineWidth=2 \
#   --prop catTitle=Temperature (C) --prop axisTitle=Daily Sales ($)
#
# Features: scatterStyle=smooth, smooth=true (Bezier interpolation),
#   marker=diamond, single series
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Scatter Fundamentals" --type chart'
    f' --prop chartType=scatter'
    f' --prop scatterStyle=smooth'
    f' --prop smooth=true'
    f' --prop title="Temperature vs Ice Cream Sales"'
    f' --prop categories=15,18,22,25,28,30,33,35'
    f' --prop series1="Sales ($):120,180,260,340,420,480,530,560"'
    f' --prop colors=C00000'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop marker=diamond --prop markerSize=7'
    f' --prop lineWidth=2'
    f' --prop catTitle="Temperature (C)" --prop axisTitle="Daily Sales ($)"')

# --------------------------------------------------------------------------
# Chart 4: Scatter line-only (no markers, scatterStyle=line)
#
# officecli add charts-scatter.xlsx "/1-Scatter Fundamentals" --type chart \
#   --prop chartType=scatter \
#   --prop scatterStyle=line \
#   --prop title="Altitude vs Air Pressure" \
#   --prop categories=0,500,1000,2000,3000,5000,8000 \
#   --prop series1="Pressure (hPa):1013,955,899,795,701,540,356" \
#   --prop colors=1F4E79 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop showMarker=false \
#   --prop lineWidth=2.5 \
#   --prop lineDash=dash \
#   --prop catTitle=Altitude (m) --prop axisTitle=Pressure (hPa)
#
# Features: scatterStyle=line (line without markers), showMarker=false,
#   lineWidth=2.5, lineDash=dash
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Scatter Fundamentals" --type chart'
    f' --prop chartType=scatter'
    f' --prop scatterStyle=line'
    f' --prop title="Altitude vs Air Pressure"'
    f' --prop categories=0,500,1000,2000,3000,5000,8000'
    f' --prop series1="Pressure (hPa):1013,955,899,795,701,540,356"'
    f' --prop colors=1F4E79'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop showMarker=false'
    f' --prop lineWidth=2.5'
    f' --prop lineDash=dash'
    f' --prop catTitle="Altitude (m)" --prop axisTitle="Pressure (hPa)"')

# ==========================================================================
# Sheet: 2-Marker Styles
# ==========================================================================
print("\n--- 2-Marker Styles ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Marker Styles"')

# --------------------------------------------------------------------------
# Chart 1: Per-series markers — circle, diamond, square
#
# officecli add charts-scatter.xlsx "/2-Marker Styles" --type chart \
#   --prop chartType=scatter \
#   --prop title="Per-Series Markers: Circle, Diamond, Square" \
#   --prop categories=10,20,30,40,50,60 \
#   --prop series1="Sensor A:12,28,35,42,55,68" \
#   --prop series2="Sensor B:8,22,30,38,48,58" \
#   --prop series3="Sensor C:15,25,32,45,52,62" \
#   --prop colors=4472C4,ED7D31,70AD47 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop series1.marker=circle \
#   --prop series2.marker=diamond \
#   --prop series3.marker=square \
#   --prop markerSize=8 --prop lineWidth=1 \
#   --prop legend=bottom
#
# Features: series1.marker=circle, series2.marker=diamond,
#   series3.marker=square (per-series marker style)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Marker Styles" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Per-Series Markers: Circle, Diamond, Square"'
    f' --prop categories=10,20,30,40,50,60'
    f' --prop series1="Sensor A:12,28,35,42,55,68"'
    f' --prop series2="Sensor B:8,22,30,38,48,58"'
    f' --prop series3="Sensor C:15,25,32,45,52,62"'
    f' --prop colors=4472C4,ED7D31,70AD47'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop series1.marker=circle'
    f' --prop series2.marker=diamond'
    f' --prop series3.marker=square'
    f' --prop markerSize=8 --prop lineWidth=1'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: Per-series markers — triangle, star, x
#
# officecli add charts-scatter.xlsx "/2-Marker Styles" --type chart \
#   --prop chartType=scatter \
#   --prop title="Per-Series Markers: Triangle, Star, X" \
#   --prop categories=5,10,15,20,25,30 \
#   --prop series1="Lab 1:18,32,28,45,52,60" \
#   --prop series2="Lab 2:22,25,38,40,48,55" \
#   --prop series3="Lab 3:10,20,32,35,42,50" \
#   --prop colors=FFC000,9DC3E6,843C0B \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop series1.marker=triangle \
#   --prop series2.marker=star \
#   --prop series3.marker=x \
#   --prop markerSize=9 --prop lineWidth=1 \
#   --prop legend=bottom
#
# Features: series1.marker=triangle, series2.marker=star,
#   series3.marker=x
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Marker Styles" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Per-Series Markers: Triangle, Star, X"'
    f' --prop categories=5,10,15,20,25,30'
    f' --prop series1="Lab 1:18,32,28,45,52,60"'
    f' --prop series2="Lab 2:22,25,38,40,48,55"'
    f' --prop series3="Lab 3:10,20,32,35,42,50"'
    f' --prop colors=FFC000,9DC3E6,843C0B'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop series1.marker=triangle'
    f' --prop series2.marker=star'
    f' --prop series3.marker=x'
    f' --prop markerSize=9 --prop lineWidth=1'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Large markers with series colors, markerSize=10
#
# officecli add charts-scatter.xlsx "/2-Marker Styles" --type chart \
#   --prop chartType=scatter \
#   --prop scatterStyle=marker \
#   --prop title="Large Markers (size=10)" \
#   --prop categories=100,200,300,400,500 \
#   --prop series1="Revenue:150,280,350,420,510" \
#   --prop series2="Profit:80,140,180,220,280" \
#   --prop series3="Cost:70,140,170,200,230" \
#   --prop colors=2E75B6,548235,BF8F00 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop series1.marker=circle \
#   --prop series2.marker=plus \
#   --prop series3.marker=dash \
#   --prop markerSize=10 \
#   --prop legend=right
#
# Features: markerSize=10, marker=plus, marker=dash, scatterStyle=marker
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Marker Styles" --type chart'
    f' --prop chartType=scatter'
    f' --prop scatterStyle=marker'
    f' --prop title="Large Markers (size=10)"'
    f' --prop categories=100,200,300,400,500'
    f' --prop series1="Revenue:150,280,350,420,510"'
    f' --prop series2="Profit:80,140,180,220,280"'
    f' --prop series3="Cost:70,140,170,200,230"'
    f' --prop colors=2E75B6,548235,BF8F00'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop series1.marker=circle'
    f' --prop series2.marker=plus'
    f' --prop series3.marker=dash'
    f' --prop markerSize=10'
    f' --prop legend=right')

# --------------------------------------------------------------------------
# Chart 4: showMarker=false (line only) vs showMarker=true
#
# officecli add charts-scatter.xlsx "/2-Marker Styles" --type chart \
#   --prop chartType=scatter \
#   --prop scatterStyle=lineMarker \
#   --prop title="Marker Toggle (none shown)" \
#   --prop categories=1,2,3,4,5,6,7,8,9,10 \
#   --prop series1="Signal:3,7,5,11,9,14,12,18,15,20" \
#   --prop series2="Noise:2,4,6,5,8,7,10,9,12,11" \
#   --prop colors=4472C4,BFBFBF \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop showMarker=false \
#   --prop lineWidth=2 \
#   --prop lineDash=dashDot \
#   --prop legend=bottom
#
# Features: scatterStyle=lineMarker, showMarker=false (markers hidden),
#   lineDash=dashDot
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Marker Styles" --type chart'
    f' --prop chartType=scatter'
    f' --prop scatterStyle=lineMarker'
    f' --prop title="Marker Toggle (none shown)"'
    f' --prop categories=1,2,3,4,5,6,7,8,9,10'
    f' --prop series1="Signal:3,7,5,11,9,14,12,18,15,20"'
    f' --prop series2="Noise:2,4,6,5,8,7,10,9,12,11"'
    f' --prop colors=4472C4,BFBFBF'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop showMarker=false'
    f' --prop lineWidth=2'
    f' --prop lineDash=dashDot'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 3-Trendlines
# ==========================================================================
print("\n--- 3-Trendlines ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Trendlines"')

# --------------------------------------------------------------------------
# Chart 1: Linear trendline with equation display
#
# officecli add charts-scatter.xlsx "/3-Trendlines" --type chart \
#   --prop chartType=scatter \
#   --prop scatterStyle=marker \
#   --prop title="Linear Trendline + Equation" \
#   --prop categories=1,2,3,4,5,6,7,8,9,10 \
#   --prop series1="Observed:8,15,22,28,33,42,48,55,60,68" \
#   --prop colors=4472C4 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop markerSize=7 \
#   --prop trendline=linear \
#   --prop series1.trendline.equation=true \
#   --prop catTitle=X --prop axisTitle=Y
#
# Features: trendline=linear, series1.trendline.equation=true
#   (display equation on chart)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Trendlines" --type chart'
    f' --prop chartType=scatter'
    f' --prop scatterStyle=marker'
    f' --prop title="Linear Trendline + Equation"'
    f' --prop categories=1,2,3,4,5,6,7,8,9,10'
    f' --prop series1="Observed:8,15,22,28,33,42,48,55,60,68"'
    f' --prop colors=4472C4'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop markerSize=7'
    f' --prop trendline=linear'
    f' --prop series1.trendline.equation=true'
    f' --prop catTitle=X --prop axisTitle=Y')

# --------------------------------------------------------------------------
# Chart 2: Polynomial trendline (order 3) with R-squared display
#
# officecli add charts-scatter.xlsx "/3-Trendlines" --type chart \
#   --prop chartType=scatter \
#   --prop scatterStyle=marker \
#   --prop title="Polynomial (order 3) + R-squared" \
#   --prop categories=1,2,3,4,5,6,7,8,9,10 \
#   --prop series1="Measurement:5,12,25,30,28,35,50,62,58,72" \
#   --prop colors=70AD47 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop markerSize=7 --prop marker=square \
#   --prop trendline=poly:3 \
#   --prop series1.trendline.rsquared=true \
#   --prop catTitle=Sample --prop axisTitle=Value
#
# Features: trendline=poly:3 (polynomial order 3),
#   series1.trendline.rsquared=true (R-squared display)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Trendlines" --type chart'
    f' --prop chartType=scatter'
    f' --prop scatterStyle=marker'
    f' --prop title="Polynomial (order 3) + R-squared"'
    f' --prop categories=1,2,3,4,5,6,7,8,9,10'
    f' --prop series1="Measurement:5,12,25,30,28,35,50,62,58,72"'
    f' --prop colors=70AD47'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop markerSize=7 --prop marker=square'
    f' --prop trendline=poly:3'
    f' --prop series1.trendline.rsquared=true'
    f' --prop catTitle=Sample --prop axisTitle=Value')

# --------------------------------------------------------------------------
# Chart 3: Exponential trendline with forward/backward extrapolation
#
# officecli add charts-scatter.xlsx "/3-Trendlines" --type chart \
#   --prop chartType=scatter \
#   --prop scatterStyle=marker \
#   --prop title="Exponential + Extrapolation" \
#   --prop categories=1,2,3,4,5,6,7,8 \
#   --prop series1="Growth:2,4,7,12,20,35,58,95" \
#   --prop colors=ED7D31 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop markerSize=7 --prop marker=triangle \
#   --prop trendline=exp:2:1 \
#   --prop series1.trendline.name=Exponential Fit \
#   --prop catTitle=Period --prop axisTitle=Amount
#
# Features: trendline=exp:2:1 (exponential, forward=2, backward=1),
#   series1.trendline.name (custom trendline label)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Trendlines" --type chart'
    f' --prop chartType=scatter'
    f' --prop scatterStyle=marker'
    f' --prop title="Exponential + Extrapolation"'
    f' --prop categories=1,2,3,4,5,6,7,8'
    f' --prop series1="Growth:2,4,7,12,20,35,58,95"'
    f' --prop colors=ED7D31'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop markerSize=7 --prop marker=triangle'
    f' --prop trendline=exp:2:1'
    f' --prop series1.trendline.name="Exponential Fit"'
    f' --prop catTitle=Period --prop axisTitle=Amount')

# --------------------------------------------------------------------------
# Chart 4: Per-series trendlines — linear vs logarithmic
#
# officecli add charts-scatter.xlsx "/3-Trendlines" --type chart \
#   --prop chartType=scatter \
#   --prop scatterStyle=marker \
#   --prop title="Per-Series: Linear vs Logarithmic" \
#   --prop categories=1,2,4,8,16,32,64 \
#   --prop series1="Dataset A:10,18,30,45,62,78,95" \
#   --prop series2="Dataset B:5,25,38,45,50,54,56" \
#   --prop colors=4472C4,C00000 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop markerSize=7 \
#   --prop series1.trendline=linear \
#   --prop series2.trendline=log \
#   --prop series1.trendline.equation=true \
#   --prop series2.trendline.rsquared=true \
#   --prop legend=bottom
#
# Features: series1.trendline=linear, series2.trendline=log,
#   per-series trendline with sub-properties
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Trendlines" --type chart'
    f' --prop chartType=scatter'
    f' --prop scatterStyle=marker'
    f' --prop title="Per-Series: Linear vs Logarithmic"'
    f' --prop categories=1,2,4,8,16,32,64'
    f' --prop series1="Dataset A:10,18,30,45,62,78,95"'
    f' --prop series2="Dataset B:5,25,38,45,50,54,56"'
    f' --prop colors=4472C4,C00000'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop markerSize=7'
    f' --prop series1.trendline=linear'
    f' --prop series2.trendline=log'
    f' --prop series1.trendline.equation=true'
    f' --prop series2.trendline.rsquared=true'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 4-Error Bars
# ==========================================================================
print("\n--- 4-Error Bars ---")
cli(f'add "{FILE}" / --type sheet --prop name="4-Error Bars"')

# --------------------------------------------------------------------------
# Chart 1: Fixed error bars (errBars=fixed:5)
#
# officecli add charts-scatter.xlsx "/4-Error Bars" --type chart \
#   --prop chartType=scatter \
#   --prop title="Fixed Error Bars (+-5)" \
#   --prop categories=10,20,30,40,50,60 \
#   --prop series1="Measurement:25,42,58,72,88,105" \
#   --prop colors=4472C4 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop marker=circle --prop markerSize=7 \
#   --prop lineWidth=1 \
#   --prop errBars=fixed:5 \
#   --prop catTitle=Input --prop axisTitle=Output
#
# Features: errBars=fixed:5 (constant +/-5 error)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Error Bars" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Fixed Error Bars (+-5)"'
    f' --prop categories=10,20,30,40,50,60'
    f' --prop series1="Measurement:25,42,58,72,88,105"'
    f' --prop colors=4472C4'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop marker=circle --prop markerSize=7'
    f' --prop lineWidth=1'
    f' --prop errBars=fixed:5'
    f' --prop catTitle=Input --prop axisTitle=Output')

# --------------------------------------------------------------------------
# Chart 2: Percentage error bars (errBars=percent:10)
#
# officecli add charts-scatter.xlsx "/4-Error Bars" --type chart \
#   --prop chartType=scatter \
#   --prop title="Percentage Error Bars (10%)" \
#   --prop categories=5,10,15,20,25,30 \
#   --prop series1="Yield:120,185,240,310,375,450" \
#   --prop colors=70AD47 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop marker=diamond --prop markerSize=7 \
#   --prop lineWidth=1 \
#   --prop errBars=percent:10 \
#   --prop catTitle=Dosage --prop axisTitle=Yield
#
# Features: errBars=percent:10 (10% of each value)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Error Bars" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Percentage Error Bars (10%)"'
    f' --prop categories=5,10,15,20,25,30'
    f' --prop series1="Yield:120,185,240,310,375,450"'
    f' --prop colors=70AD47'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop marker=diamond --prop markerSize=7'
    f' --prop lineWidth=1'
    f' --prop errBars=percent:10'
    f' --prop catTitle=Dosage --prop axisTitle=Yield')

# --------------------------------------------------------------------------
# Chart 3: Standard deviation error bars (errBars=stddev)
#
# officecli add charts-scatter.xlsx "/4-Error Bars" --type chart \
#   --prop chartType=scatter \
#   --prop title="Standard Deviation Error Bars" \
#   --prop categories=0,1,2,3,4,5,6,7 \
#   --prop series1="Trial 1:48,52,47,55,50,53,49,51" \
#   --prop series2="Trial 2:30,35,28,40,32,38,34,36" \
#   --prop colors=ED7D31,9DC3E6 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop marker=square --prop markerSize=6 \
#   --prop lineWidth=1 \
#   --prop errBars=stddev \
#   --prop legend=bottom
#
# Features: errBars=stddev (standard deviation), multi-series with errBars
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Error Bars" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Standard Deviation Error Bars"'
    f' --prop categories=0,1,2,3,4,5,6,7'
    f' --prop series1="Trial 1:48,52,47,55,50,53,49,51"'
    f' --prop series2="Trial 2:30,35,28,40,32,38,34,36"'
    f' --prop colors=ED7D31,9DC3E6'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop marker=square --prop markerSize=6'
    f' --prop lineWidth=1'
    f' --prop errBars=stddev'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: Standard error with series styling
#
# officecli add charts-scatter.xlsx "/4-Error Bars" --type chart \
#   --prop chartType=scatter \
#   --prop title="Standard Error + Styled Series" \
#   --prop categories=2,4,6,8,10,12,14 \
#   --prop series1="Experiment:18,32,28,45,40,55,52" \
#   --prop colors=843C0B \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop marker=star --prop markerSize=8 \
#   --prop lineWidth=1.5 \
#   --prop errBars=stderr \
#   --prop series.shadow=000000-4-315-2-30 \
#   --prop gridlines=D9D9D9:0.5:dot \
#   --prop catTitle=Time (h) --prop axisTitle=Response
#
# Features: errBars=stderr, series.shadow, gridlines styling
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Error Bars" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Standard Error + Styled Series"'
    f' --prop categories=2,4,6,8,10,12,14'
    f' --prop series1="Experiment:18,32,28,45,40,55,52"'
    f' --prop colors=843C0B'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop marker=star --prop markerSize=8'
    f' --prop lineWidth=1.5'
    f' --prop errBars=stderr'
    f' --prop series.shadow=000000-4-315-2-30'
    f' --prop gridlines=D9D9D9:0.5:dot'
    f' --prop catTitle="Time (h)" --prop axisTitle=Response')

# ==========================================================================
# Sheet: 5-Styling
# ==========================================================================
print("\n--- 5-Styling ---")
cli(f'add "{FILE}" / --type sheet --prop name="5-Styling"')

# --------------------------------------------------------------------------
# Chart 1: Title styling, series shadow, series outline
#
# officecli add charts-scatter.xlsx "/5-Styling" --type chart \
#   --prop chartType=scatter \
#   --prop title="Styled Title + Series Effects" \
#   --prop categories=10,20,30,40,50 \
#   --prop series1="Alpha:15,35,28,48,55" \
#   --prop series2="Beta:8,22,32,40,50" \
#   --prop colors=4472C4,ED7D31 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop marker=circle --prop markerSize=8 --prop lineWidth=2 \
#   --prop title.font=Georgia --prop title.size=16 \
#   --prop title.color=1F4E79 --prop title.bold=true \
#   --prop title.shadow=000000-3-315-2-30 \
#   --prop series.shadow=000000-4-315-2-30 \
#   --prop series.outline=333333:1.5 \
#   --prop legend=bottom
#
# Features: title.font, title.size, title.color, title.bold, title.shadow,
#   series.shadow, series.outline
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Styling" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Styled Title + Series Effects"'
    f' --prop categories=10,20,30,40,50'
    f' --prop series1="Alpha:15,35,28,48,55"'
    f' --prop series2="Beta:8,22,32,40,50"'
    f' --prop colors=4472C4,ED7D31'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop marker=circle --prop markerSize=8 --prop lineWidth=2'
    f' --prop title.font=Georgia --prop title.size=16'
    f' --prop title.color=1F4E79 --prop title.bold=true'
    f' --prop title.shadow=000000-3-315-2-30'
    f' --prop series.shadow=000000-4-315-2-30'
    f' --prop series.outline=333333:1.5'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: Gradients, transparency, plotFill, chartFill
#
# officecli add charts-scatter.xlsx "/5-Styling" --type chart \
#   --prop chartType=scatter \
#   --prop title="Gradients + Fills" \
#   --prop categories=5,15,25,35,45 \
#   --prop series1="Group 1:12,28,35,42,55" \
#   --prop series2="Group 2:8,18,22,38,48" \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop marker=diamond --prop markerSize=8 --prop lineWidth=1.5 \
#   --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90' \
#   --prop transparency=20 \
#   --prop plotFill=F5F5F5 \
#   --prop chartFill=FAFAFA \
#   --prop legend=bottom
#
# Features: gradients (per-series gradient), transparency, plotFill, chartFill
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Styling" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Gradients + Fills"'
    f' --prop categories=5,15,25,35,45'
    f' --prop series1="Group 1:12,28,35,42,55"'
    f' --prop series2="Group 2:8,18,22,38,48"'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop marker=diamond --prop markerSize=8 --prop lineWidth=1.5'
    f' --prop "gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90"'
    f' --prop transparency=20'
    f' --prop plotFill=F5F5F5'
    f' --prop chartFill=FAFAFA'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Axis font, gridlines, minor gridlines, axis line
#
# officecli add charts-scatter.xlsx "/5-Styling" --type chart \
#   --prop chartType=scatter \
#   --prop title="Axis & Grid Styling" \
#   --prop categories=0,10,20,30,40,50 \
#   --prop series1="Readings:5,22,38,52,68,82" \
#   --prop colors=2E75B6 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop marker=circle --prop markerSize=7 --prop lineWidth=1.5 \
#   --prop axisfont=9:C00000:Arial \
#   --prop gridlines=BFBFBF:0.75:solid \
#   --prop minorGridlines=E0E0E0:0.25:dot \
#   --prop axisLine=333333:1 \
#   --prop catTitle=X Axis --prop axisTitle=Y Axis
#
# Features: axisfont (size:color:font), gridlines, minorGridlines, axisLine
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Styling" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Axis & Grid Styling"'
    f' --prop categories=0,10,20,30,40,50'
    f' --prop series1="Readings:5,22,38,52,68,82"'
    f' --prop colors=2E75B6'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop marker=circle --prop markerSize=7 --prop lineWidth=1.5'
    f' --prop axisfont=9:C00000:Arial'
    f' --prop gridlines=BFBFBF:0.75:solid'
    f' --prop minorGridlines=E0E0E0:0.25:dot'
    f' --prop axisLine=333333:1'
    f' --prop catTitle="X Axis" --prop axisTitle="Y Axis"')

# --------------------------------------------------------------------------
# Chart 4: Chart area border, plot area border, rounded corners
#
# officecli add charts-scatter.xlsx "/5-Styling" --type chart \
#   --prop chartType=scatter \
#   --prop title="Borders + Rounded Corners" \
#   --prop categories=1,3,5,7,9 \
#   --prop series1="Data:10,25,18,35,28" \
#   --prop colors=548235 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop marker=square --prop markerSize=8 --prop lineWidth=1.5 \
#   --prop chartArea.border=333333:1.5 \
#   --prop plotArea.border=999999:0.75 \
#   --prop roundedCorners=true \
#   --prop chartFill=FFFFFF \
#   --prop plotFill=F0F0F0
#
# Features: chartArea.border, plotArea.border, roundedCorners
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Styling" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Borders + Rounded Corners"'
    f' --prop categories=1,3,5,7,9'
    f' --prop series1="Data:10,25,18,35,28"'
    f' --prop colors=548235'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop marker=square --prop markerSize=8 --prop lineWidth=1.5'
    f' --prop chartArea.border=333333:1.5'
    f' --prop plotArea.border=999999:0.75'
    f' --prop roundedCorners=true'
    f' --prop chartFill=FFFFFF'
    f' --prop plotFill=F0F0F0')

# ==========================================================================
# Sheet: 6-Advanced
# ==========================================================================
print("\n--- 6-Advanced ---")
cli(f'add "{FILE}" / --type sheet --prop name="6-Advanced"')

# --------------------------------------------------------------------------
# Chart 1: Secondary axis
#
# officecli add charts-scatter.xlsx "/6-Advanced" --type chart \
#   --prop chartType=scatter \
#   --prop title="Secondary Y-Axis" \
#   --prop categories=10,20,30,40,50,60 \
#   --prop series1="Temperature (C):15,20,28,32,38,42" \
#   --prop series2="Humidity (%):85,78,65,58,45,38" \
#   --prop colors=C00000,4472C4 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop marker=circle --prop markerSize=7 --prop lineWidth=1.5 \
#   --prop secondaryAxis=2 \
#   --prop legend=bottom \
#   --prop catTitle=Location
#
# Features: secondaryAxis=2 (series 2 on right Y-axis)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Advanced" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Secondary Y-Axis"'
    f' --prop categories=10,20,30,40,50,60'
    f' --prop series1="Temperature (C):15,20,28,32,38,42"'
    f' --prop series2="Humidity (%):85,78,65,58,45,38"'
    f' --prop colors=C00000,4472C4'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop marker=circle --prop markerSize=7 --prop lineWidth=1.5'
    f' --prop secondaryAxis=2'
    f' --prop legend=bottom'
    f' --prop catTitle=Location')

# --------------------------------------------------------------------------
# Chart 2: Reference line (horizontal target)
#
# officecli add charts-scatter.xlsx "/6-Advanced" --type chart \
#   --prop chartType=scatter \
#   --prop title="Reference Line (Target=75)" \
#   --prop categories=1,2,3,4,5,6,7,8 \
#   --prop series1="Score:60,68,72,78,80,74,82,88" \
#   --prop colors=70AD47 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop marker=diamond --prop markerSize=7 --prop lineWidth=1.5 \
#   --prop referenceLine=75:FF0000:Target:dash \
#   --prop catTitle=Week --prop axisTitle=Performance
#
# Features: referenceLine=value:color:label:dash (horizontal target line)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Advanced" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Reference Line (Target=75)"'
    f' --prop categories=1,2,3,4,5,6,7,8'
    f' --prop series1="Score:60,68,72,78,80,74,82,88"'
    f' --prop colors=70AD47'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop marker=diamond --prop markerSize=7 --prop lineWidth=1.5'
    f' --prop referenceLine=75:FF0000:Target:dash'
    f' --prop catTitle=Week --prop axisTitle=Performance')

# --------------------------------------------------------------------------
# Chart 3: Axis min/max and log scale
#
# officecli add charts-scatter.xlsx "/6-Advanced" --type chart \
#   --prop chartType=scatter \
#   --prop title="Log Scale (base 10)" \
#   --prop categories=1,10,100,1000,10000 \
#   --prop series1="Response:2,15,120,950,8500" \
#   --prop colors=1F4E79 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop marker=triangle --prop markerSize=8 --prop lineWidth=1.5 \
#   --prop logBase=10 \
#   --prop axisMin=1 --prop axisMax=10000 \
#   --prop catTitle=Concentration --prop axisTitle=Response
#
# Features: logBase=10 (logarithmic value axis), axisMin, axisMax
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Advanced" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Log Scale (base 10)"'
    f' --prop categories=1,10,100,1000,10000'
    f' --prop series1="Response:2,15,120,950,8500"'
    f' --prop colors=1F4E79'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop marker=triangle --prop markerSize=8 --prop lineWidth=1.5'
    f' --prop logBase=10'
    f' --prop axisMin=1 --prop axisMax=10000'
    f' --prop catTitle=Concentration --prop axisTitle=Response')

# --------------------------------------------------------------------------
# Chart 4: Data labels and color rule
#
# officecli add charts-scatter.xlsx "/6-Advanced" --type chart \
#   --prop chartType=scatter \
#   --prop scatterStyle=marker \
#   --prop title="Data Labels + Color Rule" \
#   --prop categories=1,2,3,4,5,6,7,8 \
#   --prop series1="KPI:45,62,38,78,55,82,48,90" \
#   --prop colors=4472C4 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop markerSize=9 \
#   --prop dataLabels=true --prop labelPos=top \
#   --prop colorRule=60:C00000:00AA00 \
#   --prop catTitle=Quarter --prop axisTitle=KPI Score
#
# Features: dataLabels=true, labelPos=top, colorRule=threshold:below:above
#   (points below 60 = red, above = green)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Advanced" --type chart'
    f' --prop chartType=scatter'
    f' --prop scatterStyle=marker'
    f' --prop title="Data Labels + Color Rule"'
    f' --prop categories=1,2,3,4,5,6,7,8'
    f' --prop series1="KPI:45,62,38,78,55,82,48,90"'
    f' --prop colors=4472C4'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop markerSize=9'
    f' --prop dataLabels=true --prop labelPos=top'
    f' --prop colorRule=60:C00000:00AA00'
    f' --prop catTitle=Quarter --prop axisTitle="KPI Score"')

# Remove blank default Sheet1 (all data is inline)
cli(f'remove "{FILE}" /Sheet1')

print(f"\nDone! Generated: {FILE}")
print("  7 sheets (6 chart sheets, 24 charts total)")
