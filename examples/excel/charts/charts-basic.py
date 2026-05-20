#!/usr/bin/env python3
"""
Basic Charts Showcase — column, bar, line, and area charts with all variations.

Generates: charts-basic.xlsx

Each sheet demonstrates one chart family with all its variants and key properties.
See charts-basic.md for a guide to each sheet.

Usage:
  python3 charts-basic.py
"""

import subprocess, sys, os, json, atexit

FILE = "charts-basic.xlsx"

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
# Source data — shared across all charts
# ==========================================================================
print("\n--- Populating source data ---")

data_cmds = []
for j, h in enumerate(["Month", "East", "South", "North", "West"]):
    data_cmds.append({"command": "set", "path": f"/Sheet1/{'ABCDE'[j]}1", "props": {"text": h, "bold": "true"}})

months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
east =   [120, 135, 148, 162, 155, 178, 195, 210, 188, 172, 165, 198]
south =  [95,  108, 115, 128, 142, 155, 168, 175, 160, 148, 135, 158]
north =  [88,  92,  105, 118, 125, 138, 145, 152, 140, 130, 122, 142]
west =   [110, 118, 130, 145, 138, 162, 175, 190, 170, 155, 148, 180]

for i in range(12):
    r = i + 2
    for j, val in enumerate([months[i], east[i], south[i], north[i], west[i]]):
        data_cmds.append({"command": "set", "path": f"/Sheet1/{'ABCDE'[j]}{r}", "props": {"text": str(val)}})

cli(f'batch "{FILE}" --force --commands \'{json.dumps(data_cmds)}\'')

# ==========================================================================
# Sheet: 1-Column Charts
# ==========================================================================
print("\n--- 1-Column Charts ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Column Charts"')

# --------------------------------------------------------------------------
# Chart 1: Basic clustered column from cell range with axis titles
#
# officecli add charts-basic.xlsx "/1-Column Charts" --type chart \
#   --prop chartType=column \
#   --prop title="Regional Sales by Month" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop catTitle=Month --prop axisTitle=Sales \
#   --prop axisfont=9:58626E:Arial \
#   --prop gridlines=D9D9D9:0.5:dot
#
# Features: chartType=column, dataRange, catTitle, axisTitle, axisfont, gridlines
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Column Charts" --type chart'
    f' --prop chartType=column'
    f' --prop title="Regional Sales by Month"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop catTitle=Month --prop axisTitle=Sales'
    f' --prop axisfont=9:58626E:Arial'
    f' --prop gridlines=D9D9D9:0.5:dot')

# --------------------------------------------------------------------------
# Chart 2: Stacked column with custom colors, data labels, and gap control
#
# officecli add charts-basic.xlsx "/1-Column Charts" --type chart \
#   --prop chartType=columnStacked \
#   --prop title="Stacked Regional Sales" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop colors=2E75B6,70AD47,FFC000,C00000 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelPos=center \
#   --prop gapwidth=60 \
#   --prop series.outline=FFFFFF-0.5
#
# Features: columnStacked, colors, dataLabels, labelPos, gapwidth, series.outline
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Column Charts" --type chart'
    f' --prop chartType=columnStacked'
    f' --prop title="Stacked Regional Sales"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop colors=2E75B6,70AD47,FFC000,C00000'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop dataLabels=true --prop labelPos=center'
    f' --prop gapwidth=60'
    f' --prop series.outline=FFFFFF-0.5')

# --------------------------------------------------------------------------
# Chart 3: 100% stacked column with legend position and plotFill
#
# officecli add charts-basic.xlsx "/1-Column Charts" --type chart \
#   --prop chartType=columnPercentStacked \
#   --prop title="Market Share by Month" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop legend=bottom \
#   --prop legendfont=9:8B949E \
#   --prop plotFill=F5F5F5
#
# Features: columnPercentStacked, legend=bottom, legendfont, plotFill
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Column Charts" --type chart'
    f' --prop chartType=columnPercentStacked'
    f' --prop title="Market Share by Month"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop legend=bottom'
    f' --prop legendfont=9:8B949E'
    f' --prop plotFill=F5F5F5')

# --------------------------------------------------------------------------
# Chart 4: 3D column with perspective and title styling
#
# officecli add charts-basic.xlsx "/1-Column Charts" --type chart \
#   --prop chartType=column3d \
#   --prop title="3D Regional Sales" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop view3d=15,20,30 \
#   --prop title.font=Calibri --prop title.size=16 \
#   --prop title.color=1F4E79 --prop title.bold=true
#
# Features: column3d, view3d (rotX,rotY,perspective), title.font/size/color/bold
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Column Charts" --type chart'
    f' --prop chartType=column3d'
    f' --prop title="3D Regional Sales"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop view3d=15,20,30'
    f' --prop title.font=Calibri --prop title.size=16'
    f' --prop title.color=1F4E79 --prop title.bold=true')

# ==========================================================================
# Sheet: 2-Bar Charts
# ==========================================================================
print("\n--- 2-Bar Charts ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Bar Charts"')

# --------------------------------------------------------------------------
# Chart 1: Horizontal bar with inline data and gapwidth
#
# officecli add charts-basic.xlsx "/2-Bar Charts" --type chart \
#   --prop chartType=bar \
#   --prop title="Q4 Sales by Region" \
#   --prop 'data=East:198;South:158;North:142;West:180' \
#   --prop categories=East,South,North,West \
#   --prop colors=2E75B6,70AD47,FFC000,C00000 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop gapwidth=80 \
#   --prop dataLabels=true --prop labelPos=outsideEnd
#
# Features: bar, inline data (Name:v1;Name2:v2), gapwidth, labelPos=outsideEnd
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bar Charts" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Q4 Sales by Region"'
    f' --prop "data=East:198;South:158;North:142;West:180"'
    f' --prop categories=East,South,North,West'
    f' --prop colors=2E75B6,70AD47,FFC000,C00000'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop gapwidth=80'
    f' --prop dataLabels=true --prop labelPos=outsideEnd')

# --------------------------------------------------------------------------
# Chart 2: Stacked bar with named series and overlap
#
# officecli add charts-basic.xlsx "/2-Bar Charts" --type chart \
#   --prop chartType=barStacked \
#   --prop title="H1 vs H2 Sales" \
#   --prop series1=H1:663,598,528,661 \
#   --prop series2=H2:833,718,669,868 \
#   --prop categories=East,South,North,West \
#   --prop colors=4472C4,ED7D31 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelPos=center \
#   --prop gapwidth=50 --prop overlap=0
#
# Features: barStacked, named series (series1=Name:v1,v2), overlap
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bar Charts" --type chart'
    f' --prop chartType=barStacked'
    f' --prop title="H1 vs H2 Sales"'
    f' --prop series1=H1:663,598,528,661'
    f' --prop series2=H2:833,718,669,868'
    f' --prop categories=East,South,North,West'
    f' --prop colors=4472C4,ED7D31'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop dataLabels=true --prop labelPos=center'
    f' --prop gapwidth=50 --prop overlap=0')

# --------------------------------------------------------------------------
# Chart 3: 100% stacked bar with reference line
#
# officecli add charts-basic.xlsx "/2-Bar Charts" --type chart \
#   --prop chartType=barPercentStacked \
#   --prop title="Regional Contribution %" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop referenceLine=0.5:FF0000:Target:dash \
#   --prop axisLine=333333:1:solid \
#   --prop catAxisLine=333333:1:solid
#
# Note: on a barPercentStacked chart, the value axis is 0-1 (displayed as 0%-100%),
# so a 50% reference line must be written as 0.5 — not 50.
# referenceLine supports: value | value:color | value:color:label | value:color:width:dash
# | value:color:label:dash (legacy) | value:color:width:dash:label (canonical).
# Width is in points; default 1.5pt.
#
# Features: barPercentStacked, referenceLine, axisLine, catAxisLine
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bar Charts" --type chart'
    f' --prop chartType=barPercentStacked'
    f' --prop title="Regional Contribution %"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop referenceLine=0.5:FF0000:Target:dash'
    f' --prop axisLine=333333:1:solid'
    f' --prop catAxisLine=333333:1:solid')

# --------------------------------------------------------------------------
# Chart 4: 3D bar with chart area fill and display units
#
# officecli add charts-basic.xlsx "/2-Bar Charts" --type chart \
#   --prop chartType=bar3d \
#   --prop title="3D Regional Comparison" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop view3d=10,30,20 \
#   --prop chartFill=F2F2F2 \
#   --prop style=3
#
# Features: bar3d, chartFill (chart area background), style/styleId (preset 1-48)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bar Charts" --type chart'
    f' --prop chartType=bar3d'
    f' --prop title="3D Regional Comparison"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop view3d=10,30,20'
    f' --prop chartFill=F2F2F2'
    f' --prop style=3')

# ==========================================================================
# Sheet: 3-Line Charts
# ==========================================================================
print("\n--- 3-Line Charts ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Line Charts"')

# --------------------------------------------------------------------------
# Chart 1: Line with markers and cell-range series (dotted syntax)
#
# officecli add charts-basic.xlsx "/3-Line Charts" --type chart \
#   --prop chartType=line \
#   --prop title="East Region Trend" \
#   --prop series1.name=East \
#   --prop series1.values=Sheet1!B2:B13 \
#   --prop series1.categories=Sheet1!A2:A13 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop showMarkers=true --prop marker=circle:6:2E75B6 \
#   --prop gridlines=D9D9D9:0.5:dot \
#   --prop minorGridlines=EEEEEE:0.3:dot
#
# Features: series.name/values/categories (cell range), marker (style:size:color),
#   gridlines, minorGridlines
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Line Charts" --type chart'
    f' --prop chartType=line'
    f' --prop title="East Region Trend"'
    f' --prop series1.name=East'
    f' --prop series1.values=Sheet1!B2:B13'
    f' --prop series1.categories=Sheet1!A2:A13'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop showMarkers=true --prop marker=circle:6:2E75B6'
    f' --prop gridlines=D9D9D9:0.5:dot'
    f' --prop minorGridlines=EEEEEE:0.3:dot')

# --------------------------------------------------------------------------
# Chart 2: Smooth line with custom width and no gridlines
#
# officecli add charts-basic.xlsx "/3-Line Charts" --type chart \
#   --prop chartType=line \
#   --prop title="Smoothed Sales Trend" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop smooth=true --prop lineWidth=2.5 \
#   --prop colors=0070C0,00B050,FFC000,FF0000 \
#   --prop gridlines=none \
#   --prop series.shadow=000000-4-315-2-40
#
# Features: smooth, lineWidth, gridlines=none, series.shadow (color-blur-angle-dist-opacity)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Line Charts" --type chart'
    f' --prop chartType=line'
    f' --prop title="Smoothed Sales Trend"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop smooth=true --prop lineWidth=2.5'
    f' --prop colors=0070C0,00B050,FFC000,FF0000'
    f' --prop gridlines=none'
    f' --prop series.shadow=000000-4-315-2-40')

# --------------------------------------------------------------------------
# Chart 3: Stacked line
#
# officecli add charts-basic.xlsx "/3-Line Charts" --type chart \
#   --prop chartType=lineStacked \
#   --prop title="Cumulative Sales" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop catTitle=Month --prop axisTitle=Cumulative \
#   --prop majorTickMark=outside --prop tickLabelPos=low
#
# Features: lineStacked, majorTickMark, tickLabelPos
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Line Charts" --type chart'
    f' --prop chartType=lineStacked'
    f' --prop title="Cumulative Sales"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop catTitle=Month --prop axisTitle=Cumulative'
    f' --prop majorTickMark=outside --prop tickLabelPos=low')

# --------------------------------------------------------------------------
# Chart 4: Line with dashed lines, data table, and hidden legend
#
# officecli add charts-basic.xlsx "/3-Line Charts" --type chart \
#   --prop chartType=line \
#   --prop title="Trend with Data Table" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop lineDash=dash --prop lineWidth=1.5 \
#   --prop dataTable=true \
#   --prop legend=none
#
# Features: lineDash (solid/dot/dash/dashdot/longdash), dataTable, legend=none
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Line Charts" --type chart'
    f' --prop chartType=line'
    f' --prop title="Trend with Data Table"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop lineDash=dash --prop lineWidth=1.5'
    f' --prop dataTable=true'
    f' --prop legend=none')

# ==========================================================================
# Sheet: 4-Area Charts
# ==========================================================================
print("\n--- 4-Area Charts ---")
cli(f'add "{FILE}" / --type sheet --prop name="4-Area Charts"')

# --------------------------------------------------------------------------
# Chart 1: Area with transparency and gradient fill
#
# officecli add charts-basic.xlsx "/4-Area Charts" --type chart \
#   --prop chartType=area \
#   --prop title="Sales Volume" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop transparency=40 \
#   --prop gradient=4472C4-BDD7EE:90
#
# Features: area, transparency (0-100%), gradient (color1-color2:angle)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Area Charts" --type chart'
    f' --prop chartType=area'
    f' --prop title="Sales Volume"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop transparency=40'
    f' --prop gradient=4472C4-BDD7EE:90')

# --------------------------------------------------------------------------
# Chart 2: Stacked area with plotFill and rounded corners
#
# officecli add charts-basic.xlsx "/4-Area Charts" --type chart \
#   --prop chartType=areaStacked \
#   --prop title="Stacked Volume" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop plotFill=F5F5F5 \
#   --prop roundedCorners=true \
#   --prop transparency=30
#
# Features: areaStacked, plotFill, roundedCorners
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Area Charts" --type chart'
    f' --prop chartType=areaStacked'
    f' --prop title="Stacked Volume"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop plotFill=F5F5F5'
    f' --prop roundedCorners=true'
    f' --prop transparency=30')

# --------------------------------------------------------------------------
# Chart 3: 100% stacked area with axis control
#
# officecli add charts-basic.xlsx "/4-Area Charts" --type chart \
#   --prop chartType=areaPercentStacked \
#   --prop title="Regional Mix %" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop transparency=20 \
#   --prop axisVisible=true \
#   --prop axisLine=999999:0.5:solid
#
# Features: areaPercentStacked, axisVisible, axisLine
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Area Charts" --type chart'
    f' --prop chartType=areaPercentStacked'
    f' --prop title="Regional Mix %"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop transparency=20'
    f' --prop axisVisible=true'
    f' --prop axisLine=999999:0.5:solid')

# --------------------------------------------------------------------------
# Chart 4: 3D area with perspective
#
# officecli add charts-basic.xlsx "/4-Area Charts" --type chart \
#   --prop chartType=area3d \
#   --prop title="3D Sales Volume" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop view3d=20,25,15 \
#   --prop colors=5B9BD5,A5D5A5,FFD966,F4B183
#
# Features: area3d, view3d
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Area Charts" --type chart'
    f' --prop chartType=area3d'
    f' --prop title="3D Sales Volume"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop view3d=20,25,15'
    f' --prop colors=5B9BD5,A5D5A5,FFD966,F4B183')

# ==========================================================================
# Sheet: 5-Styling
# Demonstrates all styling/layout properties on a single column chart
# ==========================================================================
print("\n--- 5-Styling ---")
cli(f'add "{FILE}" / --type sheet --prop name="5-Styling"')

# --------------------------------------------------------------------------
# Chart 1: Fully styled column chart — title, legend, axis, series effects
#
# officecli add charts-basic.xlsx "/5-Styling" --type chart \
#   --prop chartType=column \
#   --prop title="Fully Styled Chart" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=0 --prop y=0 --prop width=14 --prop height=20 \
#   --prop title.font=Georgia --prop title.size=18 \
#   --prop title.color=1F4E79 --prop title.bold=true \
#   --prop title.shadow=000000-3-315-2-30 \
#   --prop legendfont=10:444444:Helvetica \
#   --prop legend=right \
#   --prop axisfont=9:58626E:Arial \
#   --prop catTitle=Month --prop axisTitle=Revenue \
#   --prop gridlines=CCCCCC:0.5:dot \
#   --prop plotFill=FAFAFA \
#   --prop chartFill=FFFFFF \
#   --prop series.outline=FFFFFF-0.5 \
#   --prop series.shadow=000000-3-315-2-25 \
#   --prop gapwidth=100 \
#   --prop roundedCorners=true \
#   --prop referenceLine=160:FF0000:1:dash \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000
#
# Features: title.font/size/color/bold/shadow, legendfont, axisfont,
#   series.outline, series.shadow, roundedCorners, referenceLine
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Styling" --type chart'
    f' --prop chartType=column'
    f' --prop title="Fully Styled Chart"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=0 --prop y=0 --prop width=14 --prop height=20'
    f' --prop title.font=Georgia --prop title.size=18'
    f' --prop title.color=1F4E79 --prop title.bold=true'
    f' --prop title.shadow=000000-3-315-2-30'
    f' --prop legendfont=10:444444:Helvetica'
    f' --prop legend=right'
    f' --prop axisfont=9:58626E:Arial'
    f' --prop catTitle=Month --prop axisTitle=Revenue'
    f' --prop gridlines=CCCCCC:0.5:dot'
    f' --prop plotFill=FAFAFA'
    f' --prop chartFill=FFFFFF'
    f' --prop series.outline=FFFFFF-0.5'
    f' --prop series.shadow=000000-3-315-2-25'
    f' --prop gapwidth=100'
    f' --prop roundedCorners=true'
    f' --prop referenceLine=160:FF0000:1:dash'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000')

# --------------------------------------------------------------------------
# Chart 2: Column with secondary axis (dual Y-axis)
#
# officecli add charts-basic.xlsx "/5-Styling" --type chart \
#   --prop chartType=column \
#   --prop title="Sales vs Growth Rate" \
#   --prop series1=Sales:120,135,148,162 \
#   --prop series2=Growth:5.2,8.1,12.3,15.6 \
#   --prop categories=Q1,Q2,Q3,Q4 \
#   --prop x=15 --prop y=0 --prop width=10 --prop height=20 \
#   --prop secondaryAxis=2 \
#   --prop colors=4472C4,FF0000
#
# Features: secondaryAxis (comma-separated 1-based series indices for second Y-axis)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Styling" --type chart'
    f' --prop chartType=column'
    f' --prop title="Sales vs Growth Rate"'
    f' --prop series1=Sales:120,135,148,162'
    f' --prop series2=Growth:5.2,8.1,12.3,15.6'
    f' --prop categories=Q1,Q2,Q3,Q4'
    f' --prop x=15 --prop y=0 --prop width=10 --prop height=20'
    f' --prop secondaryAxis=2'
    f' --prop colors=4472C4,FF0000')

# --------------------------------------------------------------------------
# Chart 3: Column with individual point colors and inverted negatives
#
# officecli add charts-basic.xlsx "/5-Styling" --type chart \
#   --prop chartType=column \
#   --prop title="Quarterly P&L" \
#   --prop series1=P&L:500,300,-200,800 \
#   --prop categories=Q1,Q2,Q3,Q4 \
#   --prop x=0 --prop y=21 --prop width=10 --prop height=18 \
#   --prop point1.color=70AD47 --prop point2.color=70AD47 \
#   --prop point3.color=FF0000 --prop point4.color=70AD47 \
#   --prop invertIfNeg=true \
#   --prop dataLabels=true --prop labelPos=outsideEnd
#
# Features: point{N}.color (per-point coloring), invertIfNeg
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Styling" --type chart'
    f' --prop chartType=column'
    f' --prop title="Quarterly P&L"'
    f' --prop "series1=P&L:500,300,-200,800"'
    f' --prop categories=Q1,Q2,Q3,Q4'
    f' --prop x=0 --prop y=21 --prop width=10 --prop height=18'
    f' --prop point1.color=70AD47 --prop point2.color=70AD47'
    f' --prop point3.color=FF0000 --prop point4.color=70AD47'
    f' --prop invertIfNeg=true'
    f' --prop dataLabels=true --prop labelPos=outsideEnd')

# --------------------------------------------------------------------------
# Chart 4: Line with gradient plot area and custom data labels
#
# officecli add charts-basic.xlsx "/5-Styling" --type chart \
#   --prop chartType=line \
#   --prop title="Custom Labels Demo" \
#   --prop series1=Revenue:100,200,300,250 \
#   --prop categories=Q1,Q2,Q3,Q4 \
#   --prop x=11 --prop y=21 --prop width=14 --prop height=18 \
#   --prop plotFill=E8F0FE-FFFFFF:90 \
#   --prop showMarkers=true --prop marker=diamond:8:4472C4 \
#   --prop lineWidth=2 \
#   --prop dataLabels=true --prop labelPos=top \
#   --prop dataLabels.numFmt=#,##0 \
#   --prop dataLabel3.text=Peak!
#
# Features: plotFill gradient (color1-color2:angle), marker styles (diamond),
#   dataLabels.numFmt, dataLabel{N}.text (custom text for one label)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Styling" --type chart'
    f' --prop chartType=line'
    f' --prop title="Custom Labels Demo"'
    f' --prop series1=Revenue:100,200,300,250'
    f' --prop categories=Q1,Q2,Q3,Q4'
    f' --prop x=11 --prop y=21 --prop width=14 --prop height=18'
    f' --prop plotFill=E8F0FE-FFFFFF:90'
    f' --prop showMarkers=true --prop marker=diamond:8:4472C4'
    f' --prop lineWidth=2'
    f' --prop dataLabels=true --prop labelPos=top'
    f' --prop dataLabels.numFmt=#,##0'
    f' --prop dataLabel3.text=Peak!')

# ==========================================================================
# Sheet: 6-Layout
# Manual layout of plot area, title, legend; axis orientation; log scale;
# display units; label font and separator; error bars
# ==========================================================================
print("\n--- 6-Layout ---")
cli(f'add "{FILE}" / --type sheet --prop name="6-Layout"')

# --------------------------------------------------------------------------
# Chart 1: Manual layout positioning of plot area, title, legend
#
# officecli add charts-basic.xlsx "/6-Layout" --type chart \
#   --prop chartType=column \
#   --prop title="Manual Layout" \
#   --prop dataRange=Sheet1!A1:C13 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop plotArea.x=0.15 --prop plotArea.y=0.15 \
#   --prop plotArea.w=0.7 --prop plotArea.h=0.7 \
#   --prop title.x=0.3 --prop title.y=0.01 \
#   --prop legend.x=0.02 --prop legend.y=0.4 \
#   --prop legend.overlay=true
#
# Features: plotArea.x/y/w/h (0-1 fraction), title.x/y, legend.x/y, legend.overlay
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Layout" --type chart'
    f' --prop chartType=column'
    f' --prop title="Manual Layout"'
    f' --prop dataRange=Sheet1!A1:C13'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop plotArea.x=0.15 --prop plotArea.y=0.15'
    f' --prop plotArea.w=0.7 --prop plotArea.h=0.7'
    f' --prop title.x=0.3 --prop title.y=0.01'
    f' --prop legend.x=0.02 --prop legend.y=0.4'
    f' --prop legend.overlay=true')

# --------------------------------------------------------------------------
# Chart 2: Reversed axis, log scale, display units
#
# officecli add charts-basic.xlsx "/6-Layout" --type chart \
#   --prop chartType=bar \
#   --prop title="Log Scale + Reversed Axis" \
#   --prop series1=Revenue:10,100,1000,10000 \
#   --prop categories=Startup,Small,Medium,Enterprise \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop logBase=10 \
#   --prop axisOrientation=maxMin \
#   --prop dispUnits=thousands
#
# Features: logBase (logarithmic scale), axisOrientation=maxMin (reversed),
#   dispUnits (thousands/millions)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Layout" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Log Scale + Reversed Axis"'
    f' --prop series1=Revenue:10,100,1000,10000'
    f' --prop categories=Startup,Small,Medium,Enterprise'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop logBase=10'
    f' --prop axisOrientation=maxMin'
    f' --prop dispUnits=thousands')

# --------------------------------------------------------------------------
# Chart 3: Label font, separator, leader lines, and per-label layout
#
# officecli add charts-basic.xlsx "/6-Layout" --type chart \
#   --prop chartType=column \
#   --prop title="Label Formatting" \
#   --prop series1=Sales:120,200,150,180 \
#   --prop categories=Q1,Q2,Q3,Q4 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelPos=outsideEnd \
#   --prop labelFont=11:2E75B6:true \
#   --prop dataLabels.separator=": " \
#   --prop dataLabel2.text=Best! \
#   --prop dataLabel3.delete=true
#
# Features: labelFont (size:color:bold), dataLabels.separator,
#   dataLabel{N}.text (custom), dataLabel{N}.delete (hide one label)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Layout" --type chart'
    f' --prop chartType=column'
    f' --prop title="Label Formatting"'
    f' --prop series1=Sales:120,200,150,180'
    f' --prop categories=Q1,Q2,Q3,Q4'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop dataLabels=true --prop labelPos=outsideEnd'
    f' --prop labelFont=11:2E75B6:true'
    f' --prop "dataLabels.separator=: "'
    f' --prop dataLabel2.text=Best!'
    f' --prop dataLabel3.delete=true')

# --------------------------------------------------------------------------
# Chart 4: Error bars, minor ticks, opacity
#
# officecli add charts-basic.xlsx "/6-Layout" --type chart \
#   --prop chartType=line \
#   --prop title="Error Bars + Ticks" \
#   --prop series1=Measurement:50,55,48,62,58 \
#   --prop categories=Mon,Tue,Wed,Thu,Fri \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop showMarkers=true --prop marker=square:7:4472C4 \
#   --prop errBars=percentage \
#   --prop majorTickMark=outside --prop minorTickMark=inside \
#   --prop opacity=80
#
# Features: errBars (percentage/stdDev/fixed), minorTickMark, opacity (0-100%)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Layout" --type chart'
    f' --prop chartType=line'
    f' --prop title="Error Bars + Ticks"'
    f' --prop series1=Measurement:50,55,48,62,58'
    f' --prop categories=Mon,Tue,Wed,Thu,Fri'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop showMarkers=true --prop marker=square:7:4472C4'
    f' --prop errBars=percentage'
    f' --prop majorTickMark=outside --prop minorTickMark=inside'
    f' --prop opacity=80')

# ==========================================================================
# Sheet: 7-Effects
# Gradients, conditional color, area fill, title glow, preset themes
# ==========================================================================
print("\n--- 7-Effects ---")
cli(f'add "{FILE}" / --type sheet --prop name="7-Effects"')

# --------------------------------------------------------------------------
# Chart 1: Per-series gradients
#
# officecli add charts-basic.xlsx "/7-Effects" --type chart \
#   --prop chartType=column \
#   --prop title="Per-Series Gradients" \
#   --prop series1=East:120,135,148 \
#   --prop series2=West:110,118,130 \
#   --prop categories=Q1,Q2,Q3 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90'
#
# Features: gradients (per-series, semicolon-separated "C1-C2:angle")
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/7-Effects" --type chart'
    f' --prop chartType=column'
    f' --prop title="Per-Series Gradients"'
    f' --prop series1=East:120,135,148'
    f' --prop series2=West:110,118,130'
    f' --prop categories=Q1,Q2,Q3'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop "gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90"')

# --------------------------------------------------------------------------
# Chart 2: Area fill gradient and title glow effect
#
# officecli add charts-basic.xlsx "/7-Effects" --type chart \
#   --prop chartType=area \
#   --prop title="Glow Title + Area Fill" \
#   --prop dataRange=Sheet1!A1:C13 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop areafill=4472C4-BDD7EE:90 \
#   --prop transparency=30 \
#   --prop title.glow=4472C4-8-60 \
#   --prop title.size=16
#
# Features: areafill (area gradient), title.glow (color-radius-opacity)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/7-Effects" --type chart'
    f' --prop chartType=area'
    f' --prop title="Glow Title + Area Fill"'
    f' --prop dataRange=Sheet1!A1:C13'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop areafill=4472C4-BDD7EE:90'
    f' --prop transparency=30'
    f' --prop title.glow=4472C4-8-60'
    f' --prop title.size=16')

# --------------------------------------------------------------------------
# Chart 3: Conditional coloring rule
#
# officecli add charts-basic.xlsx "/7-Effects" --type chart \
#   --prop chartType=column \
#   --prop title="Conditional Colors" \
#   --prop series1=Score:85,42,91,38,76,55 \
#   --prop categories=A,B,C,D,E,F \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop colorRule=60:FF0000:70AD47 \
#   --prop dataLabels=true --prop labelPos=outsideEnd
#
# Features: colorRule (threshold:belowColor:aboveColor — values below 60 red, above green)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/7-Effects" --type chart'
    f' --prop chartType=column'
    f' --prop title="Conditional Colors"'
    f' --prop series1=Score:85,42,91,38,76,55'
    f' --prop categories=A,B,C,D,E,F'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop colorRule=60:FF0000:70AD47'
    f' --prop dataLabels=true --prop labelPos=outsideEnd')

# --------------------------------------------------------------------------
# Chart 4: Preset style/theme and leader lines
#
# officecli add charts-basic.xlsx "/7-Effects" --type chart \
#   --prop chartType=column \
#   --prop title="Preset Style 26" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop style=26 \
#   --prop dataLabels=true \
#   --prop dataLabels.showLeaderLines=true
#
# Features: style (preset 1-48), dataLabels.showLeaderLines
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/7-Effects" --type chart'
    f' --prop chartType=column'
    f' --prop title="Preset Style 26"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop style=26'
    f' --prop dataLabels=true'
    f' --prop dataLabels.showLeaderLines=true')

print(f"\nDone! Generated: {FILE}")
print("  8 sheets (Sheet1 data + 7 chart sheets, 28 charts total)")
