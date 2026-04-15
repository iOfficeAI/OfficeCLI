#!/usr/bin/env python3
"""
Advanced Charts Showcase — scatter, bubble, combo, radar, and stock charts.

Generates: charts-advanced.xlsx

Usage:
  python3 charts-advanced.py
"""

import subprocess, sys, os, json, atexit

FILE = "charts-advanced.xlsx"

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
# Sheet: 1-Scatter & Bubble
# ==========================================================================
print("\n--- 1-Scatter & Bubble ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Scatter & Bubble"')

# --------------------------------------------------------------------------
# Chart 1: Scatter with markers — circle markers, line connecting points
#
# officecli add charts-advanced.xlsx "/1-Scatter & Bubble" --type chart \
#   --prop chartType=scatter \
#   --prop title="Scatter: Markers & Line" \
#   --prop categories=1,2,3,4,5,6 \
#   --prop series1="SeriesA:10,25,15,40,30,50" \
#   --prop series2="SeriesB:5,18,22,35,28,42" \
#   --prop colors=4472C4,ED7D31 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop marker=circle --prop markerSize=8 \
#   --prop lineWidth=1.5 \
#   --prop legend=bottom
#
# Features: chartType=scatter, categories as X values, marker=circle,
#   markerSize, lineWidth, legend=bottom
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Scatter & Bubble" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Scatter: Markers & Line"'
    f' --prop categories=1,2,3,4,5,6'
    f' --prop series1=SeriesA:10,25,15,40,30,50'
    f' --prop series2=SeriesB:5,18,22,35,28,42'
    f' --prop colors=4472C4,ED7D31'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop marker=circle --prop markerSize=8'
    f' --prop lineWidth=1.5'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: Scatter with smooth curve and trendline (reference line)
#
# officecli add charts-advanced.xlsx "/1-Scatter & Bubble" --type chart \
#   --prop chartType=scatter \
#   --prop title="Scatter: Smooth + Trendline" \
#   --prop categories=1,2,3,4,5,6,7,8 \
#   --prop series1="Growth:3,7,12,20,28,35,40,45" \
#   --prop colors=70AD47 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop smooth=true \
#   --prop marker=diamond --prop markerSize=7 \
#   --prop referenceLine=25:FF0000:Target:dash \
#   --prop axisTitle=Value --prop catTitle=Period
#
# Features: smooth=true (smooth curve), marker=diamond,
#   referenceLine (trendline overlay), axisTitle, catTitle
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Scatter & Bubble" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Scatter: Smooth + Trendline"'
    f' --prop categories=1,2,3,4,5,6,7,8'
    f' --prop series1=Growth:3,7,12,20,28,35,40,45'
    f' --prop colors=70AD47'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop smooth=true'
    f' --prop marker=diamond --prop markerSize=7'
    f' --prop referenceLine=25:FF0000:Target:dash'
    f' --prop axisTitle=Value --prop catTitle=Period')

# --------------------------------------------------------------------------
# Chart 3: Scatter with varied marker styles per series
#
# officecli add charts-advanced.xlsx "/1-Scatter & Bubble" --type chart \
#   --prop chartType=scatter \
#   --prop title="Scatter: Marker Styles" \
#   --prop categories=10,20,30,40,50 \
#   --prop series1="Squares:8,22,18,35,30" \
#   --prop series2="Triangles:15,10,28,20,42" \
#   --prop series3="Stars:5,30,12,45,25" \
#   --prop colors=4472C4,ED7D31,70AD47 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop series1.marker=square \
#   --prop series2.marker=triangle \
#   --prop series3.marker=star \
#   --prop markerSize=9 \
#   --prop lineWidth=1 \
#   --prop gridlines=D9D9D9:0.5:dot
#
# Features: per-series marker style (series{N}.marker), gridlines styling
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Scatter & Bubble" --type chart'
    f' --prop chartType=scatter'
    f' --prop title="Scatter: Marker Styles"'
    f' --prop categories=10,20,30,40,50'
    f' --prop series1=Squares:8,22,18,35,30'
    f' --prop series2=Triangles:15,10,28,20,42'
    f' --prop series3=Stars:5,30,12,45,25'
    f' --prop colors=4472C4,ED7D31,70AD47'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop series1.marker=square'
    f' --prop series2.marker=triangle'
    f' --prop series3.marker=star'
    f' --prop markerSize=9'
    f' --prop lineWidth=1'
    f' --prop gridlines=D9D9D9:0.5:dot')

# --------------------------------------------------------------------------
# Chart 4: Bubble chart with size data
#
# officecli add charts-advanced.xlsx "/1-Scatter & Bubble" --type chart \
#   --prop chartType=bubble \
#   --prop title="Bubble: Market Size" \
#   --prop categories=10,25,40,60,80 \
#   --prop series1="ProductA:30,50,20,70,45" \
#   --prop series2="ProductB:15,35,55,40,60" \
#   --prop colors=4472C4,ED7D31 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop bubbleScale=80 \
#   --prop legend=right \
#   --prop dataLabels=false
#
# Features: chartType=bubble, categories as X, series as Y values,
#   bubble sizes default to Y values, bubbleScale to control sizing
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Scatter & Bubble" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Bubble: Market Size"'
    f' --prop categories=10,25,40,60,80'
    f' --prop series1=ProductA:30,50,20,70,45'
    f' --prop series2=ProductB:15,35,55,40,60'
    f' --prop colors=4472C4,ED7D31'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop bubbleScale=80'
    f' --prop legend=right')

# ==========================================================================
# Sheet: 2-Combo & Radar
# ==========================================================================
print("\n--- 2-Combo & Radar ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Combo & Radar"')

# --------------------------------------------------------------------------
# Chart 1: Combo chart — bar+line with comboSplit
#
# officecli add charts-advanced.xlsx "/2-Combo & Radar" --type chart \
#   --prop chartType=combo \
#   --prop title="Combo: Sales (Bar) + Growth % (Line)" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop series1="Revenue:120,145,132,168,155,180" \
#   --prop series2="Expenses:80,92,85,98,90,105" \
#   --prop series3="Growth:8,12,6,15,10,16" \
#   --prop colors=4472C4,ED7D31,70AD47 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop comboSplit=2 \
#   --prop legend=bottom \
#   --prop axisTitle=Amount --prop catTitle=Month
#
# Features: chartType=combo, comboSplit=2 (first 2 series as bars,
#   remaining as lines), categories as X labels
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Combo & Radar" --type chart'
    f' --prop chartType=combo'
    f' --prop title="Combo: Sales (Bar) + Growth % (Line)"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop series1=Revenue:120,145,132,168,155,180'
    f' --prop series2=Expenses:80,92,85,98,90,105'
    f' --prop series3=Growth:8,12,6,15,10,16'
    f' --prop colors=4472C4,ED7D31,70AD47'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop comboSplit=2'
    f' --prop legend=bottom'
    f' --prop axisTitle=Amount --prop catTitle=Month')

# --------------------------------------------------------------------------
# Chart 2: Combo with secondary axis
#
# officecli add charts-advanced.xlsx "/2-Combo & Radar" --type chart \
#   --prop chartType=combo \
#   --prop title="Combo: Volume (Bar) + Price (Line, 2nd Axis)" \
#   --prop categories=Q1,Q2,Q3,Q4 \
#   --prop series1="Volume:1200,1450,1320,1680" \
#   --prop series2="AvgPrice:45,52,48,58" \
#   --prop colors=5B9BD5,FF0000 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop comboSplit=1 \
#   --prop secondaryAxis=2 \
#   --prop legend=bottom
#
# Features: comboSplit=1, secondaryAxis=2 (series 2 on right Y-axis)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Combo & Radar" --type chart'
    f' --prop chartType=combo'
    f' --prop title="Combo: Volume (Bar) + Price (Line, 2nd Axis)"'
    f' --prop categories=Q1,Q2,Q3,Q4'
    f' --prop series1=Volume:1200,1450,1320,1680'
    f' --prop series2=AvgPrice:45,52,48,58'
    f' --prop colors=5B9BD5,FF0000'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop comboSplit=1'
    f' --prop secondaryAxis=2'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Combo with combotypes — per-series type control
#
# officecli add charts-advanced.xlsx "/2-Combo & Radar" --type chart \
#   --prop chartType=combo \
#   --prop title="Combo: Mixed Types (combotypes)" \
#   --prop categories=A,B,C,D,E \
#   --prop series1="Bars:30,45,28,52,40" \
#   --prop series2="MoreBars:20,30,22,38,28" \
#   --prop series3="Lines:12,18,15,22,16" \
#   --prop series4="Area:8,12,10,15,11" \
#   --prop colors=4472C4,5B9BD5,ED7D31,70AD47 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop combotypes=column,column,line,area \
#   --prop legend=bottom
#
# Features: combotypes (per-series type: column, column, line, area)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Combo & Radar" --type chart'
    f' --prop chartType=combo'
    f' --prop title="Combo: Mixed Types (combotypes)"'
    f' --prop categories=A,B,C,D,E'
    f' --prop series1=Bars:30,45,28,52,40'
    f' --prop series2=MoreBars:20,30,22,38,28'
    f' --prop series3=Lines:12,18,15,22,16'
    f' --prop series4=Area:8,12,10,15,11'
    f' --prop colors=4472C4,5B9BD5,ED7D31,70AD47'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop combotypes=column,column,line,area'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: Radar (spider) chart with multiple series
#
# officecli add charts-advanced.xlsx "/2-Combo & Radar" --type chart \
#   --prop chartType=radar \
#   --prop title="Radar: Skills Comparison" \
#   --prop categories=Speed,Strength,Stamina,Agility,Accuracy \
#   --prop series1="AthleteA:80,65,90,75,85" \
#   --prop series2="AthleteB:70,85,60,90,70" \
#   --prop series3="AthleteC:90,70,75,65,80" \
#   --prop colors=4472C4,ED7D31,70AD47 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop radarStyle=marker \
#   --prop legend=bottom
#
# Features: chartType=radar, categories as spoke labels,
#   multiple series, radarStyle=marker
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Combo & Radar" --type chart'
    f' --prop chartType=radar'
    f' --prop title="Radar: Skills Comparison"'
    f' --prop categories=Speed,Strength,Stamina,Agility,Accuracy'
    f' --prop series1=AthleteA:80,65,90,75,85'
    f' --prop series2=AthleteB:70,85,60,90,70'
    f' --prop series3=AthleteC:90,70,75,65,80'
    f' --prop colors=4472C4,ED7D31,70AD47'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop radarStyle=marker'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 3-Stock & More Radar
# ==========================================================================
print("\n--- 3-Stock & More Radar ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Stock & Radar"')

# --------------------------------------------------------------------------
# Chart 1: Stock (OHLC) chart — Open-High-Low-Close
#
# officecli add charts-advanced.xlsx "/3-Stock & Radar" --type chart \
#   --prop chartType=stock \
#   --prop title="Stock: OHLC Daily Prices" \
#   --prop categories=Mon,Tue,Wed,Thu,Fri \
#   --prop series1="Open:145,148,150,147,152" \
#   --prop series2="High:152,155,157,153,160" \
#   --prop series3="Low:143,146,148,144,150" \
#   --prop series4="Close:148,150,147,152,158" \
#   --prop x=0 --prop y=0 --prop width=14 --prop height=18 \
#   --prop legend=bottom \
#   --prop catTitle=Day --prop axisTitle=Price
#
# Features: chartType=stock, 4 series (Open/High/Low/Close),
#   categories as date labels, catTitle, axisTitle
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Stock & Radar" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Stock: OHLC Daily Prices"'
    f' --prop categories=Mon,Tue,Wed,Thu,Fri'
    f' --prop series1=Open:145,148,150,147,152'
    f' --prop series2=High:152,155,157,153,160'
    f' --prop series3=Low:143,146,148,144,150'
    f' --prop series4=Close:148,150,147,152,158'
    f' --prop x=0 --prop y=0 --prop width=14 --prop height=18'
    f' --prop legend=bottom'
    f' --prop catTitle=Day --prop axisTitle=Price')

# --------------------------------------------------------------------------
# Chart 2: Stock chart — weekly OHLC with date categories
#
# officecli add charts-advanced.xlsx "/3-Stock & Radar" --type chart \
#   --prop chartType=stock \
#   --prop title="Stock: Weekly OHLC (6 Weeks)" \
#   --prop categories=W1,W2,W3,W4,W5,W6 \
#   --prop series1="Open:100,104,102,108,105,110" \
#   --prop series2="High:106,110,108,115,112,118" \
#   --prop series3="Low:98,101,100,105,103,107" \
#   --prop series4="Close:104,102,108,105,110,115" \
#   --prop x=15 --prop y=0 --prop width=14 --prop height=18 \
#   --prop gridlines=E0E0E0:0.75 \
#   --prop legend=bottom
#
# Features: stock chart with 6 weeks of OHLC, gridlines styling
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Stock & Radar" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Stock: Weekly OHLC (6 Weeks)"'
    f' --prop categories=W1,W2,W3,W4,W5,W6'
    f' --prop series1=Open:100,104,102,108,105,110'
    f' --prop series2=High:106,110,108,115,112,118'
    f' --prop series3=Low:98,101,100,105,103,107'
    f' --prop series4=Close:104,102,108,105,110,115'
    f' --prop x=15 --prop y=0 --prop width=14 --prop height=18'
    f' --prop gridlines=E0E0E0:0.75'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Radar — filled style (spider web)
#
# officecli add charts-advanced.xlsx "/3-Stock & Radar" --type chart \
#   --prop chartType=radar \
#   --prop title="Radar: Product Ratings (Filled)" \
#   --prop categories=Quality,Price,Design,Support,Delivery \
#   --prop series1="BrandX:85,70,90,75,80" \
#   --prop series2="BrandY:70,90,65,85,75" \
#   --prop colors=4472C4,70AD47 \
#   --prop x=0 --prop y=19 --prop width=14 --prop height=18 \
#   --prop radarStyle=filled \
#   --prop transparency=40 \
#   --prop legend=bottom
#
# Features: radarStyle=filled, transparency (fill alpha), multiple series
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Stock & Radar" --type chart'
    f' --prop chartType=radar'
    f' --prop title="Radar: Product Ratings (Filled)"'
    f' --prop categories=Quality,Price,Design,Support,Delivery'
    f' --prop series1=BrandX:85,70,90,75,80'
    f' --prop series2=BrandY:70,90,65,85,75'
    f' --prop colors=4472C4,70AD47'
    f' --prop x=0 --prop y=19 --prop width=14 --prop height=18'
    f' --prop radarStyle=filled'
    f' --prop transparency=40'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: Bubble — single series with explicit large differences in size
#
# officecli add charts-advanced.xlsx "/3-Stock & Radar" --type chart \
#   --prop chartType=bubble \
#   --prop title="Bubble: Regional Opportunity" \
#   --prop categories=5,15,30,50,70,90 \
#   --prop series1="Regions:20,45,30,80,55,65" \
#   --prop colors=4472C4 \
#   --prop x=15 --prop y=19 --prop width=14 --prop height=18 \
#   --prop bubbleScale=100 \
#   --prop legend=none \
#   --prop axisTitle=Revenue --prop catTitle=Market Size
#
# Features: bubble with single series, bubbleScale=100, legend=none,
#   axisTitle and catTitle labels
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Stock & Radar" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Bubble: Regional Opportunity"'
    f' --prop categories=5,15,30,50,70,90'
    f' --prop series1=Regions:20,45,30,80,55,65'
    f' --prop colors=4472C4'
    f' --prop x=15 --prop y=19 --prop width=14 --prop height=18'
    f' --prop bubbleScale=100'
    f' --prop legend=none'
    f' --prop axisTitle=Revenue --prop catTitle=Market Size')

print(f"\nDone! Generated: {FILE}")
print("  3 sheets (1-Scatter & Bubble, 2-Combo & Radar, 3-Stock & Radar)")
print("  12 charts total: scatter(3), bubble(2), combo(3), radar(2), stock(2)")
