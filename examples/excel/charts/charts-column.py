#!/usr/bin/env python3
"""
Column & Bar Charts Showcase — column, columnStacked, columnPercentStacked, and column3d with all variations.

Generates: charts-column.xlsx

Every column chart feature officecli supports is demonstrated at least once:
gap width, overlap, bar shapes, axis scaling, gridlines, data labels,
legend positioning, reference lines, secondary axis, gradients,
transparency, shadows, manual layout, and 3D rotation.

7 sheets, 28 charts total.

  1-Column Fundamentals   4 charts — data input variants, axis titles, inline/cell-range/data
  2-Column Variants       4 charts — columnStacked, columnPercentStacked, column3d
  3-Column Styling        4 charts — title styling, series effects, gradients, transparency
  4-Axis & Gridlines      4 charts — axis scaling, log scale, reverse, display units
  5-Labels & Legend       4 charts — data labels, custom labels, legend layout
  6-Effects & Advanced    4 charts — secondary axis, reference line, glow/shadow, colorRule
  7-Bar Shape & Gap       4 charts — gapwidth, overlap, 3D shapes (cylinder, cone, pyramid)

Usage:
  python3 charts-column.py
"""

import subprocess, sys, os, json, atexit

FILE = "charts-column.xlsx"

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
# Sheet: 1-Column Fundamentals
# ==========================================================================
print("\n--- 1-Column Fundamentals ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Column Fundamentals"')

# --------------------------------------------------------------------------
# Chart 1: Basic column with dataRange and axis titles
#
# officecli add charts-column.xlsx "/1-Column Fundamentals" --type chart \
#   --prop chartType=column \
#   --prop title="Monthly Sales by Region" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop catTitle=Month --prop axisTitle=Revenue \
#   --prop axisfont=9:58626E:Arial \
#   --prop gridlines=D9D9D9:0.5:dot \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000
#
# Features: chartType=column, dataRange, catTitle, axisTitle, axisfont,
#   gridlines, colors
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Column Fundamentals" --type chart'
    f' --prop chartType=column'
    f' --prop title="Monthly Sales by Region"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop catTitle=Month --prop axisTitle=Revenue'
    f' --prop axisfont=9:58626E:Arial'
    f' --prop gridlines=D9D9D9:0.5:dot'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000')

# --------------------------------------------------------------------------
# Chart 2: Inline series with custom colors and gap width
#
# officecli add charts-column.xlsx "/1-Column Fundamentals" --type chart \
#   --prop chartType=column \
#   --prop title="Q1 Product Sales" \
#   --prop series1="Laptops:320,280,350,310" \
#   --prop series2="Phones:450,420,480,460" \
#   --prop series3="Tablets:180,160,200,190" \
#   --prop categories=Jan,Feb,Mar,Apr \
#   --prop colors=2E75B6,C00000,70AD47 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop gapwidth=80 \
#   --prop legend=bottom
#
# Features: inline series (series1=Name:v1,v2,...), colors, gapwidth,
#   legend=bottom
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Column Fundamentals" --type chart'
    f' --prop chartType=column'
    f' --prop title="Q1 Product Sales"'
    f' --prop series1="Laptops:320,280,350,310"'
    f' --prop series2="Phones:450,420,480,460"'
    f' --prop series3="Tablets:180,160,200,190"'
    f' --prop categories=Jan,Feb,Mar,Apr'
    f' --prop colors=2E75B6,C00000,70AD47'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop gapwidth=80'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Dotted syntax with cell ranges
#
# officecli add charts-column.xlsx "/1-Column Fundamentals" --type chart \
#   --prop chartType=column \
#   --prop title="East vs South (Cell Range)" \
#   --prop series1.name=East \
#   --prop series1.values=Sheet1!B2:B13 \
#   --prop series1.categories=Sheet1!A2:A13 \
#   --prop series2.name=South \
#   --prop series2.values=Sheet1!C2:C13 \
#   --prop series2.categories=Sheet1!A2:A13 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop colors=4472C4,ED7D31 \
#   --prop gridlines=D9D9D9:0.5:dot \
#   --prop minorGridlines=EEEEEE:0.3:dot
#
# Features: series.name/values/categories (cell range via dotted syntax),
#   minorGridlines
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Column Fundamentals" --type chart'
    f' --prop chartType=column'
    f' --prop title="East vs South (Cell Range)"'
    f' --prop series1.name=East'
    f' --prop series1.values=Sheet1!B2:B13'
    f' --prop series1.categories=Sheet1!A2:A13'
    f' --prop series2.name=South'
    f' --prop series2.values=Sheet1!C2:C13'
    f' --prop series2.categories=Sheet1!A2:A13'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop colors=4472C4,ED7D31'
    f' --prop gridlines=D9D9D9:0.5:dot'
    f' --prop minorGridlines=EEEEEE:0.3:dot')

# --------------------------------------------------------------------------
# Chart 4: data= shorthand format
#
# officecli add charts-column.xlsx "/1-Column Fundamentals" --type chart \
#   --prop chartType=column \
#   --prop title="Weekly Output" \
#   --prop 'data=Team A:85,92,78,95,88;Team B:70,80,85,90,75' \
#   --prop categories=Mon,Tue,Wed,Thu,Fri \
#   --prop colors=0070C0,FF6600 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop legend=right
#
# Features: data (inline shorthand Name:v1;Name2:v2), legend=right
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Column Fundamentals" --type chart'
    f' --prop chartType=column'
    f' --prop title="Weekly Output"'
    f' --prop "data=Team A:85,92,78,95,88;Team B:70,80,85,90,75"'
    f' --prop categories=Mon,Tue,Wed,Thu,Fri'
    f' --prop colors=0070C0,FF6600'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop legend=right')

# ==========================================================================
# Sheet: 2-Column Variants
# ==========================================================================
print("\n--- 2-Column Variants ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Column Variants"')

# --------------------------------------------------------------------------
# Chart 1: Stacked column with center data labels and series outline
#
# officecli add charts-column.xlsx "/2-Column Variants" --type chart \
#   --prop chartType=columnStacked \
#   --prop title="Stacked Sales by Region" \
#   --prop dataRange=Sheet1!A1:E7 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000 \
#   --prop dataLabels=center \
#   --prop series.outline=FFFFFF-0.5 \
#   --prop legend=bottom
#
# Features: columnStacked, dataLabels=center, series.outline
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Column Variants" --type chart'
    f' --prop chartType=columnStacked'
    f' --prop title="Stacked Sales by Region"'
    f' --prop dataRange=Sheet1!A1:E7'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000'
    f' --prop dataLabels=center'
    f' --prop series.outline=FFFFFF-0.5'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: 100% stacked column with axis number format
#
# officecli add charts-column.xlsx "/2-Column Variants" --type chart \
#   --prop chartType=columnPercentStacked \
#   --prop title="Regional Contribution %" \
#   --prop dataRange=Sheet1!A1:E7 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop colors=1F4E79,2E75B6,9DC3E6,BDD7EE \
#   --prop axisNumFmt=0% \
#   --prop legend=bottom \
#   --prop gridlines=E0E0E0:0.5:solid
#
# Features: columnPercentStacked, axisNumFmt=0%, legend=bottom
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Column Variants" --type chart'
    f' --prop chartType=columnPercentStacked'
    f' --prop title="Regional Contribution %"'
    f' --prop dataRange=Sheet1!A1:E7'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop colors=1F4E79,2E75B6,9DC3E6,BDD7EE'
    f' --prop axisNumFmt=0%'
    f' --prop legend=bottom'
    f' --prop gridlines=E0E0E0:0.5:solid')

# --------------------------------------------------------------------------
# Chart 3: 3D column with perspective and style
#
# officecli add charts-column.xlsx "/2-Column Variants" --type chart \
#   --prop chartType=column3d \
#   --prop title="3D Regional Trends" \
#   --prop dataRange=Sheet1!A1:E7 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop view3d=15,20,30 \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000 \
#   --prop chartFill=F8F8F8 \
#   --prop style=3
#
# Features: column3d, view3d (rotX,rotY,perspective), style (preset 1-48)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Column Variants" --type chart'
    f' --prop chartType=column3d'
    f' --prop title="3D Regional Trends"'
    f' --prop dataRange=Sheet1!A1:E7'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop view3d=15,20,30'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000'
    f' --prop chartFill=F8F8F8'
    f' --prop style=3')

# --------------------------------------------------------------------------
# Chart 4: 3D stacked column with gap depth
#
# officecli add charts-column.xlsx "/2-Column Variants" --type chart \
#   --prop chartType=column3d \
#   --prop title="3D Stacked with Gap Depth" \
#   --prop series1="East:120,135,148,162,155,178" \
#   --prop series2="South:95,108,115,128,142,155" \
#   --prop series3="North:88,92,105,118,125,138" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop view3d=15,20,30 \
#   --prop gapDepth=200 \
#   --prop colors=2E75B6,ED7D31,70AD47 \
#   --prop legend=right
#
# Features: column3d stacked, gapDepth=200 (3D depth spacing)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Column Variants" --type chart'
    f' --prop chartType=column3d'
    f' --prop title="3D Stacked with Gap Depth"'
    f' --prop "series1=East:120,135,148,162,155,178"'
    f' --prop "series2=South:95,108,115,128,142,155"'
    f' --prop "series3=North:88,92,105,118,125,138"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop view3d=15,20,30'
    f' --prop gapDepth=200'
    f' --prop colors=2E75B6,ED7D31,70AD47'
    f' --prop legend=right')

# ==========================================================================
# Sheet: 3-Column Styling
# ==========================================================================
print("\n--- 3-Column Styling ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Column Styling"')

# --------------------------------------------------------------------------
# Chart 1: Title styling — font, size, color, bold
#
# officecli add charts-column.xlsx "/3-Column Styling" --type chart \
#   --prop chartType=column \
#   --prop title="Styled Title Demo" \
#   --prop dataRange=Sheet1!A1:E7 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop title.font=Georgia --prop title.size=16 \
#   --prop title.color=1F4E79 --prop title.bold=true \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000 \
#   --prop legend=bottom
#
# Features: title.font=Georgia, title.size=16, title.color=1F4E79,
#   title.bold=true
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Column Styling" --type chart'
    f' --prop chartType=column'
    f' --prop title="Styled Title Demo"'
    f' --prop dataRange=Sheet1!A1:E7'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop title.font=Georgia --prop title.size=16'
    f' --prop title.color=1F4E79 --prop title.bold=true'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: Series shadow and outline effects
#
# officecli add charts-column.xlsx "/3-Column Styling" --type chart \
#   --prop chartType=column \
#   --prop title="Shadow & Outline Effects" \
#   --prop series1="Revenue:320,280,350,310,340" \
#   --prop series2="Cost:210,195,230,220,215" \
#   --prop categories=Q1,Q2,Q3,Q4,Q5 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop colors=4472C4,C00000 \
#   --prop series.shadow=000000-4-315-2-40 \
#   --prop series.outline=FFFFFF-0.5 \
#   --prop gapwidth=100 \
#   --prop legend=bottom
#
# Features: series.shadow (color-blur-angle-dist-opacity),
#   series.outline (color-width)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Column Styling" --type chart'
    f' --prop chartType=column'
    f' --prop title="Shadow & Outline Effects"'
    f' --prop "series1=Revenue:320,280,350,310,340"'
    f' --prop "series2=Cost:210,195,230,220,215"'
    f' --prop categories=Q1,Q2,Q3,Q4,Q5'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop colors=4472C4,C00000'
    f' --prop series.shadow=000000-4-315-2-40'
    f' --prop series.outline=FFFFFF-0.5'
    f' --prop gapwidth=100'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Per-series gradient fills
#
# officecli add charts-column.xlsx "/3-Column Styling" --type chart \
#   --prop chartType=column \
#   --prop title="Gradient Columns" \
#   --prop series1="East:120,135,148,162" \
#   --prop series2="South:95,108,115,128" \
#   --prop series3="North:88,92,105,118" \
#   --prop categories=Q1,Q2,Q3,Q4 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90;70AD47-C5E0B4:90' \
#   --prop legend=bottom
#
# Features: gradients (per-series gradient fills, start-end:angle)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Column Styling" --type chart'
    f' --prop chartType=column'
    f' --prop title="Gradient Columns"'
    f' --prop "series1=East:120,135,148,162"'
    f' --prop "series2=South:95,108,115,128"'
    f' --prop "series3=North:88,92,105,118"'
    f' --prop categories=Q1,Q2,Q3,Q4'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop "gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90;70AD47-C5E0B4:90"'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: Transparency + plotFill gradient + chartFill + roundedCorners
#
# officecli add charts-column.xlsx "/3-Column Styling" --type chart \
#   --prop chartType=column \
#   --prop title="Transparent Columns on Gradient" \
#   --prop dataRange=Sheet1!A1:E7 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop transparency=30 \
#   --prop plotFill=F0F4F8-D6E4F0:90 \
#   --prop chartFill=FFFFFF \
#   --prop colors=1F4E79,2E75B6,5B9BD5,9DC3E6 \
#   --prop roundedCorners=true \
#   --prop legend=bottom
#
# Features: transparency=30, plotFill gradient, chartFill, roundedCorners
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Column Styling" --type chart'
    f' --prop chartType=column'
    f' --prop title="Transparent Columns on Gradient"'
    f' --prop dataRange=Sheet1!A1:E7'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop transparency=30'
    f' --prop plotFill=F0F4F8-D6E4F0:90'
    f' --prop chartFill=FFFFFF'
    f' --prop colors=1F4E79,2E75B6,5B9BD5,9DC3E6'
    f' --prop roundedCorners=true'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 4-Axis & Gridlines
# ==========================================================================
print("\n--- 4-Axis & Gridlines ---")
cli(f'add "{FILE}" / --type sheet --prop name="4-Axis & Gridlines"')

# --------------------------------------------------------------------------
# Chart 1: Custom axis scaling — min, max, majorUnit, minorUnit
#
# officecli add charts-column.xlsx "/4-Axis & Gridlines" --type chart \
#   --prop chartType=column \
#   --prop title="Custom Axis Scale (50–250)" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop axisMin=50 --prop axisMax=250 --prop majorUnit=50 \
#   --prop minorUnit=25 \
#   --prop gridlines=D0D0D0:0.5:solid \
#   --prop minorGridlines=EEEEEE:0.3:dot \
#   --prop axisLine=C00000:1.5:solid \
#   --prop catAxisLine=2E75B6:1.5:solid \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000
#
# Features: axisMin, axisMax, majorUnit, minorUnit,
#   axisLine (value axis line styling), catAxisLine (category axis line)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Axis & Gridlines" --type chart'
    f' --prop chartType=column'
    f' --prop title="Custom Axis Scale (50-250)"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop axisMin=50 --prop axisMax=250 --prop majorUnit=50'
    f' --prop minorUnit=25'
    f' --prop gridlines=D0D0D0:0.5:solid'
    f' --prop minorGridlines=EEEEEE:0.3:dot'
    f' --prop axisLine=C00000:1.5:solid'
    f' --prop catAxisLine=2E75B6:1.5:solid'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000')

# --------------------------------------------------------------------------
# Chart 2: Logarithmic scale with reversed axis
#
# officecli add charts-column.xlsx "/4-Axis & Gridlines" --type chart \
#   --prop chartType=column \
#   --prop title="Log Scale (Base 10)" \
#   --prop series1="Growth:1,10,100,1000,5000" \
#   --prop categories=Year 1,Year 2,Year 3,Year 4,Year 5 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop logBase=10 \
#   --prop axisReverse=true \
#   --prop colors=C00000 \
#   --prop axisTitle="Value (log)" \
#   --prop catTitle=Year \
#   --prop gridlines=E0E0E0:0.5:dash
#
# Features: logBase=10 (logarithmic scale), axisReverse=true
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Axis & Gridlines" --type chart'
    f' --prop chartType=column'
    f' --prop title="Log Scale (Base 10)"'
    f' --prop "series1=Growth:1,10,100,1000,5000"'
    f' --prop "categories=Year 1,Year 2,Year 3,Year 4,Year 5"'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop logBase=10'
    f' --prop axisReverse=true'
    f' --prop colors=C00000'
    f' --prop axisTitle="Value (log)"'
    f' --prop catTitle=Year'
    f' --prop gridlines=E0E0E0:0.5:dash')

# --------------------------------------------------------------------------
# Chart 3: Display units and axis number format
#
# officecli add charts-column.xlsx "/4-Axis & Gridlines" --type chart \
#   --prop chartType=column \
#   --prop title="Revenue (in Thousands)" \
#   --prop series1="Revenue:12000,18500,22000,31000,45000,52000" \
#   --prop series2="Cost:8000,11000,14000,19500,28000,33000" \
#   --prop categories=2020,2021,2022,2023,2024,2025 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop dispUnits=thousands \
#   --prop axisNumFmt=#,##0 \
#   --prop colors=2E75B6,C00000 \
#   --prop catTitle=Year --prop axisTitle=Amount (K) \
#   --prop majorTickMark=outside --prop minorTickMark=inside \
#   --prop legend=bottom
#
# Features: dispUnits=thousands, axisNumFmt=#,##0,
#   majorTickMark=outside, minorTickMark=inside
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Axis & Gridlines" --type chart'
    f' --prop chartType=column'
    f' --prop title="Revenue (in Thousands)"'
    f' --prop "series1=Revenue:12000,18500,22000,31000,45000,52000"'
    f' --prop "series2=Cost:8000,11000,14000,19500,28000,33000"'
    f' --prop categories=2020,2021,2022,2023,2024,2025'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop dispUnits=thousands'
    f' --prop axisNumFmt=#,##0'
    f' --prop colors=2E75B6,C00000'
    f' --prop catTitle=Year --prop axisTitle="Amount (K)"'
    f' --prop majorTickMark=outside --prop minorTickMark=inside'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: Hidden axes with data table
#
# officecli add charts-column.xlsx "/4-Axis & Gridlines" --type chart \
#   --prop chartType=column \
#   --prop title="Minimal Chart with Data Table" \
#   --prop dataRange=Sheet1!A1:E7 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop gridlines=none \
#   --prop axisVisible=false \
#   --prop dataTable=true \
#   --prop legend=none \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000
#
# Features: gridlines=none, axisVisible=false, dataTable=true, legend=none
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Axis & Gridlines" --type chart'
    f' --prop chartType=column'
    f' --prop title="Minimal Chart with Data Table"'
    f' --prop dataRange=Sheet1!A1:E7'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop gridlines=none'
    f' --prop axisVisible=false'
    f' --prop dataTable=true'
    f' --prop legend=none'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000')

# ==========================================================================
# Sheet: 5-Labels & Legend
# ==========================================================================
print("\n--- 5-Labels & Legend ---")
cli(f'add "{FILE}" / --type sheet --prop name="5-Labels & Legend"')

# --------------------------------------------------------------------------
# Chart 1: Data labels with number format and styled label font
#
# officecli add charts-column.xlsx "/5-Labels & Legend" --type chart \
#   --prop chartType=column \
#   --prop title="Sales with Labels" \
#   --prop series1="Revenue:120,180,210,250,280" \
#   --prop categories=Jan,Feb,Mar,Apr,May \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop colors=4472C4 \
#   --prop dataLabels=true --prop labelPos=outsideEnd \
#   --prop labelFont=9:333333:true \
#   --prop dataLabels.numFmt=#,##0
#
# Features: dataLabels=true, labelPos=outsideEnd, labelFont (size:color:bold),
#   dataLabels.numFmt
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Labels & Legend" --type chart'
    f' --prop chartType=column'
    f' --prop title="Sales with Labels"'
    f' --prop "series1=Revenue:120,180,210,250,280"'
    f' --prop categories=Jan,Feb,Mar,Apr,May'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop colors=4472C4'
    f' --prop dataLabels=true --prop labelPos=outsideEnd'
    f' --prop labelFont=9:333333:true'
    f' --prop dataLabels.numFmt=#,##0')

# --------------------------------------------------------------------------
# Chart 2: Custom individual labels — delete some, highlight peak
#
# officecli add charts-column.xlsx "/5-Labels & Legend" --type chart \
#   --prop chartType=column \
#   --prop title="Peak Highlight" \
#   --prop series1="Sales:88,120,165,210,195,178" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop colors=2E75B6 \
#   --prop dataLabels=true --prop labelPos=outsideEnd \
#   --prop dataLabel1.delete=true --prop dataLabel2.delete=true \
#   --prop dataLabel3.delete=true \
#   --prop point4.color=C00000 \
#   --prop dataLabel4.text=Peak! \
#   --prop dataLabel5.delete=true --prop dataLabel6.delete=true
#
# Features: dataLabel{N}.delete, dataLabel{N}.text, point{N}.color
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Labels & Legend" --type chart'
    f' --prop chartType=column'
    f' --prop title="Peak Highlight"'
    f' --prop "series1=Sales:88,120,165,210,195,178"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop colors=2E75B6'
    f' --prop dataLabels=true --prop labelPos=outsideEnd'
    f' --prop dataLabel1.delete=true --prop dataLabel2.delete=true'
    f' --prop dataLabel3.delete=true'
    f' --prop point4.color=C00000'
    f' --prop dataLabel4.text=Peak!'
    f' --prop dataLabel5.delete=true --prop dataLabel6.delete=true')

# --------------------------------------------------------------------------
# Chart 3: Legend positioning and overlay with styled legend font
#
# officecli add charts-column.xlsx "/5-Labels & Legend" --type chart \
#   --prop chartType=column \
#   --prop title="Legend Overlay on Chart" \
#   --prop dataRange=Sheet1!A1:E7 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000 \
#   --prop legend=right \
#   --prop legend.overlay=true \
#   --prop legendfont=10:333333:Calibri \
#   --prop plotFill=F5F5F5
#
# Features: legend=right, legend.overlay=true, legendfont (size:color:fontname),
#   plotFill
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Labels & Legend" --type chart'
    f' --prop chartType=column'
    f' --prop title="Legend Overlay on Chart"'
    f' --prop dataRange=Sheet1!A1:E7'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000'
    f' --prop legend=right'
    f' --prop legend.overlay=true'
    f' --prop legendfont=10:333333:Calibri'
    f' --prop plotFill=F5F5F5')

# --------------------------------------------------------------------------
# Chart 4: Manual layout — plotArea, title, and legend positioning
#
# officecli add charts-column.xlsx "/5-Labels & Legend" --type chart \
#   --prop chartType=column \
#   --prop title="Manual Layout Control" \
#   --prop dataRange=Sheet1!A1:E7 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop colors=2E75B6,ED7D31,70AD47,FFC000 \
#   --prop plotArea.x=0.12 --prop plotArea.y=0.18 \
#   --prop plotArea.w=0.82 --prop plotArea.h=0.55 \
#   --prop title.x=0.25 --prop title.y=0.02 \
#   --prop legend.x=0.15 --prop legend.y=0.82 \
#   --prop legend.w=0.7 --prop legend.h=0.12 \
#   --prop title.font=Arial --prop title.size=13 \
#   --prop title.bold=true
#
# Features: plotArea.x/y/w/h, title.x/y, legend.x/y/w/h (manual layout)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Labels & Legend" --type chart'
    f' --prop chartType=column'
    f' --prop title="Manual Layout Control"'
    f' --prop dataRange=Sheet1!A1:E7'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop colors=2E75B6,ED7D31,70AD47,FFC000'
    f' --prop plotArea.x=0.12 --prop plotArea.y=0.18'
    f' --prop plotArea.w=0.82 --prop plotArea.h=0.55'
    f' --prop title.x=0.25 --prop title.y=0.02'
    f' --prop legend.x=0.15 --prop legend.y=0.82'
    f' --prop legend.w=0.7 --prop legend.h=0.12'
    f' --prop title.font=Arial --prop title.size=13'
    f' --prop title.bold=true')

# ==========================================================================
# Sheet: 6-Effects & Advanced
# ==========================================================================
print("\n--- 6-Effects & Advanced ---")
cli(f'add "{FILE}" / --type sheet --prop name="6-Effects & Advanced"')

# --------------------------------------------------------------------------
# Chart 1: Secondary axis — dual Y-axis
#
# officecli add charts-column.xlsx "/6-Effects & Advanced" --type chart \
#   --prop chartType=column \
#   --prop title="Revenue vs Growth Rate" \
#   --prop series1="Revenue:120,180,250,310,380,420" \
#   --prop series2="Growth %:50,33,39,24,23,11" \
#   --prop categories=2020,2021,2022,2023,2024,2025 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop secondaryAxis=2 \
#   --prop colors=2E75B6,C00000 \
#   --prop catTitle=Year --prop axisTitle=Revenue \
#   --prop dataLabels=true --prop labelPos=outsideEnd \
#   --prop legend=bottom
#
# Features: secondaryAxis=2 (series 2 on right-hand axis)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Effects & Advanced" --type chart'
    f' --prop chartType=column'
    f' --prop title="Revenue vs Growth Rate"'
    f' --prop "series1=Revenue:120,180,250,310,380,420"'
    f' --prop "series2=Growth %:50,33,39,24,23,11"'
    f' --prop categories=2020,2021,2022,2023,2024,2025'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop secondaryAxis=2'
    f' --prop colors=2E75B6,C00000'
    f' --prop catTitle=Year --prop axisTitle=Revenue'
    f' --prop dataLabels=true --prop labelPos=outsideEnd'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: Reference line (target/threshold)
#
# officecli add charts-column.xlsx "/6-Effects & Advanced" --type chart \
#   --prop chartType=column \
#   --prop title="vs Target (150)" \
#   --prop dataRange=Sheet1!A1:C13 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop colors=4472C4,70AD47 \
#   --prop referenceLine=150:FF0000:1.5:dash \
#   --prop legend=bottom
#
# referenceLine format: value:color:width:dash
#
# Features: referenceLine (horizontal target line)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Effects & Advanced" --type chart'
    f' --prop chartType=column'
    f' --prop title="vs Target (150)"'
    f' --prop dataRange=Sheet1!A1:C13'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop colors=4472C4,70AD47'
    f' --prop referenceLine=150:FF0000:1.5:dash'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Title glow and shadow effects
#
# officecli add charts-column.xlsx "/6-Effects & Advanced" --type chart \
#   --prop chartType=column \
#   --prop title="Glow & Shadow Effects" \
#   --prop series1="East:120,135,148,162,155,178" \
#   --prop series2="West:110,118,130,145,138,162" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop colors=4472C4,ED7D31 \
#   --prop title.glow=4472C4-8-60 \
#   --prop title.shadow=000000-3-315-2-40 \
#   --prop title.font=Calibri --prop title.size=16 \
#   --prop title.bold=true --prop title.color=1F4E79 \
#   --prop series.shadow=000000-3-315-1-30 \
#   --prop plotFill=F0F4F8 --prop chartFill=FFFFFF
#
# Features: title.glow (color-radius-opacity), title.shadow,
#   series.shadow on column charts
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Effects & Advanced" --type chart'
    f' --prop chartType=column'
    f' --prop title="Glow & Shadow Effects"'
    f' --prop "series1=East:120,135,148,162,155,178"'
    f' --prop "series2=West:110,118,130,145,138,162"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop colors=4472C4,ED7D31'
    f' --prop title.glow=4472C4-8-60'
    f' --prop title.shadow=000000-3-315-2-40'
    f' --prop title.font=Calibri --prop title.size=16'
    f' --prop title.bold=true --prop title.color=1F4E79'
    f' --prop series.shadow=000000-3-315-1-30'
    f' --prop plotFill=F0F4F8 --prop chartFill=FFFFFF')

# --------------------------------------------------------------------------
# Chart 4: Conditional coloring with chart/plot borders
#
# officecli add charts-column.xlsx "/6-Effects & Advanced" --type chart \
#   --prop chartType=column \
#   --prop title="Profit: Conditional Colors" \
#   --prop series1="Profit:80,120,-30,160,-50,200,140,-20,180,90" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop colors=2E75B6 \
#   --prop colorRule=0:C00000:70AD47 \
#   --prop referenceLine=0:888888:1:solid \
#   --prop chartArea.border=D0D0D0:1:solid \
#   --prop plotArea.border=E0E0E0:0.5:dot \
#   --prop dataLabels=true --prop labelPos=outsideEnd \
#   --prop labelFont=8:666666:false
#
# colorRule format: threshold:belowColor:aboveColor
#
# Features: colorRule (threshold-based conditional coloring),
#   chartArea.border, plotArea.border
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Effects & Advanced" --type chart'
    f' --prop chartType=column'
    f' --prop title="Profit: Conditional Colors"'
    f' --prop "series1=Profit:80,120,-30,160,-50,200,140,-20,180,90"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop colors=2E75B6'
    f' --prop colorRule=0:C00000:70AD47'
    f' --prop referenceLine=0:888888:1:solid'
    f' --prop chartArea.border=D0D0D0:1:solid'
    f' --prop plotArea.border=E0E0E0:0.5:dot'
    f' --prop dataLabels=true --prop labelPos=outsideEnd'
    f' --prop labelFont=8:666666:false')

# ==========================================================================
# Sheet: 7-Bar Shape & Gap
# ==========================================================================
print("\n--- 7-Bar Shape & Gap ---")
cli(f'add "{FILE}" / --type sheet --prop name="7-Bar Shape & Gap"')

# --------------------------------------------------------------------------
# Chart 1: Narrow gap width (bars close together)
#
# officecli add charts-column.xlsx "/7-Bar Shape & Gap" --type chart \
#   --prop chartType=column \
#   --prop title="Narrow Gap (30%)" \
#   --prop dataRange=Sheet1!A1:E7 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop gapwidth=30 \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000 \
#   --prop legend=bottom
#
# Features: gapwidth=30 (narrow gaps between column groups)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/7-Bar Shape & Gap" --type chart'
    f' --prop chartType=column'
    f' --prop title="Narrow Gap (30%)"'
    f' --prop dataRange=Sheet1!A1:E7'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop gapwidth=30'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: Wide gap with negative overlap (separated bars within group)
#
# officecli add charts-column.xlsx "/7-Bar Shape & Gap" --type chart \
#   --prop chartType=column \
#   --prop title="Wide Gap + Negative Overlap" \
#   --prop dataRange=Sheet1!A1:E7 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop gapwidth=200 \
#   --prop overlap=-50 \
#   --prop colors=2E75B6,ED7D31,70AD47,FFC000 \
#   --prop legend=bottom
#
# Features: gapwidth=200 (wide gap), overlap=-50 (negative = bars separated)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/7-Bar Shape & Gap" --type chart'
    f' --prop chartType=column'
    f' --prop title="Wide Gap + Negative Overlap"'
    f' --prop dataRange=Sheet1!A1:E7'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop gapwidth=200'
    f' --prop overlap=-50'
    f' --prop colors=2E75B6,ED7D31,70AD47,FFC000'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: 3D column with cylinder shape
#
# officecli add charts-column.xlsx "/7-Bar Shape & Gap" --type chart \
#   --prop chartType=column3d \
#   --prop title="Cylinder Shape" \
#   --prop series1="East:120,135,148,162,155,178" \
#   --prop series2="South:95,108,115,128,142,155" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop shape=cylinder \
#   --prop view3d=15,20,30 \
#   --prop colors=4472C4,ED7D31 \
#   --prop legend=bottom
#
# Features: shape=cylinder (3D column bar shape)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/7-Bar Shape & Gap" --type chart'
    f' --prop chartType=column3d'
    f' --prop title="Cylinder Shape"'
    f' --prop "series1=East:120,135,148,162,155,178"'
    f' --prop "series2=South:95,108,115,128,142,155"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop shape=cylinder'
    f' --prop view3d=15,20,30'
    f' --prop colors=4472C4,ED7D31'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: 3D column with cone/pyramid shapes
#
# officecli add charts-column.xlsx "/7-Bar Shape & Gap" --type chart \
#   --prop chartType=column3d \
#   --prop title="Cone Shape" \
#   --prop series1="North:88,92,105,118,125,138" \
#   --prop series2="West:110,118,130,145,138,162" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop shape=cone \
#   --prop view3d=15,20,30 \
#   --prop colors=70AD47,FFC000 \
#   --prop legend=bottom
#
# Features: shape=cone (3D column bar shape — also supports pyramid)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/7-Bar Shape & Gap" --type chart'
    f' --prop chartType=column3d'
    f' --prop title="Cone Shape"'
    f' --prop "series1=North:88,92,105,118,125,138"'
    f' --prop "series2=West:110,118,130,145,138,162"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop shape=cone'
    f' --prop view3d=15,20,30'
    f' --prop colors=70AD47,FFC000'
    f' --prop legend=bottom')

print(f"\nDone! Generated: {FILE}")
print("  8 sheets (Sheet1 data + 7 chart sheets, 28 charts total)")
