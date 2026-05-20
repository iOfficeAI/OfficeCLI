#!/usr/bin/env python3
"""
Area Charts Showcase — area, areaStacked, areaPercentStacked, and area3d with all variations.

Generates: charts-area.xlsx

Every area chart feature officecli supports is demonstrated at least once:
area fills, gradients, transparency, stacking, axis scaling, gridlines,
data labels, legend positioning, reference lines, secondary axis,
shadows, manual layout, and 3D rotation.

5 sheets, 20 charts total.

  1-Area Fundamentals     4 charts — data input variants, transparency, area fills, gradients
  2-Area Variants         4 charts — areaStacked, areaPercentStacked, area3d
  3-Area Styling          4 charts — title styling, shadows, gridlines, chart/plot fills
  4-Labels & Legend       4 charts — data labels, per-point colors, legend, manual layout
  5-Advanced              4 charts — secondary axis, reference line, axis scaling, effects

Usage:
  python3 charts-area.py
"""

import subprocess, sys, os, json, atexit

FILE = "charts-area.xlsx"

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
for j, h in enumerate(["Month", "Organic", "Paid", "Social", "Referral"]):
    data_cmds.append({"command": "set", "path": f"/Sheet1/{'ABCDE'[j]}1", "props": {"text": h, "bold": "true"}})

months   = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
organic  = [4200, 4800, 5100, 5600, 6200, 6800, 7500, 8100, 7600, 7200, 6900, 7800]
paid     = [3100, 3500, 3800, 4200, 4800, 5200, 5800, 6300, 5900, 5500, 5100, 5700]
social   = [1800, 2100, 2400, 2800, 3200, 3600, 4000, 4300, 3900, 3500, 3200, 3800]
referral = [1200, 1400, 1500, 1700, 1900, 2100, 2300, 2500, 2300, 2100, 1900, 2200]

for i in range(12):
    r = i + 2
    for j, val in enumerate([months[i], organic[i], paid[i], social[i], referral[i]]):
        data_cmds.append({"command": "set", "path": f"/Sheet1/{'ABCDE'[j]}{r}", "props": {"text": str(val)}})

cli(f'batch "{FILE}" --force --commands \'{json.dumps(data_cmds)}\'')

# ==========================================================================
# Sheet: 1-Area Fundamentals
# ==========================================================================
print("\n--- 1-Area Fundamentals ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Area Fundamentals"')

# --------------------------------------------------------------------------
# Chart 1: Basic area chart with dataRange, axis titles, and custom colors
#
# officecli add charts-area.xlsx "/1-Area Fundamentals" --type chart \
#   --prop chartType=area \
#   --prop title="Website Traffic Overview" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop catTitle=Month --prop axisTitle=Visitors \
#   --prop gridlines=D9D9D9:0.5:dot
#
# Features: chartType=area, dataRange, colors, catTitle, axisTitle, gridlines
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Area Fundamentals" --type chart'
    f' --prop chartType=area'
    f' --prop title="Website Traffic Overview"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop catTitle=Month --prop axisTitle=Visitors'
    f' --prop gridlines=D9D9D9:0.5:dot')

# --------------------------------------------------------------------------
# Chart 2: Inline series with transparency
#
# officecli add charts-area.xlsx "/1-Area Fundamentals" --type chart \
#   --prop chartType=area \
#   --prop title="Quarterly Revenue Streams" \
#   --prop series1="Subscriptions:120,180,210,250" \
#   --prop series2="One-time:90,140,160,200" \
#   --prop series3="Services:60,85,110,145" \
#   --prop categories=Q1,Q2,Q3,Q4 \
#   --prop colors=2E75B6,70AD47,FFC000 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop transparency=40 \
#   --prop legend=bottom
#
# Features: inline series, transparency (0-100), legend=bottom
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Area Fundamentals" --type chart'
    f' --prop chartType=area'
    f' --prop title="Quarterly Revenue Streams"'
    f' --prop series1="Subscriptions:120,180,210,250"'
    f' --prop series2="One-time:90,140,160,200"'
    f' --prop series3="Services:60,85,110,145"'
    f' --prop categories=Q1,Q2,Q3,Q4'
    f' --prop colors=2E75B6,70AD47,FFC000'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop transparency=40'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Area with areafill gradient
#
# officecli add charts-area.xlsx "/1-Area Fundamentals" --type chart \
#   --prop chartType=area \
#   --prop title="Monthly Active Users" \
#   --prop series1="Users:3200,3800,4500,5100,5800,6400" \
#   --prop categories=Jul,Aug,Sep,Oct,Nov,Dec \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop areafill=4472C4-BDD7EE:90 \
#   --prop legend=none
#
# Features: areafill (gradient from-to:angle), legend=none, single series
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Area Fundamentals" --type chart'
    f' --prop chartType=area'
    f' --prop title="Monthly Active Users"'
    f' --prop series1="Users:3200,3800,4500,5100,5800,6400"'
    f' --prop categories=Jul,Aug,Sep,Oct,Nov,Dec'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop areafill=4472C4-BDD7EE:90'
    f' --prop legend=none')

# --------------------------------------------------------------------------
# Chart 4: Per-series gradient fills
#
# officecli add charts-area.xlsx "/1-Area Fundamentals" --type chart \
#   --prop chartType=area \
#   --prop title="Revenue by Channel" \
#   --prop series1="Direct:45,52,61,70" \
#   --prop series2="Partner:30,38,42,55" \
#   --prop categories=Q1,Q2,Q3,Q4 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90' \
#   --prop legend=right --prop legendfont=10:333333:Calibri
#
# Features: gradients (per-series gradient fills from-to:angle;...),
#   legendfont (size:color:font)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Area Fundamentals" --type chart'
    f' --prop chartType=area'
    f' --prop title="Revenue by Channel"'
    f' --prop series1="Direct:45,52,61,70"'
    f' --prop series2="Partner:30,38,42,55"'
    f' --prop categories=Q1,Q2,Q3,Q4'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop "gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90"'
    f' --prop legend=right --prop legendfont=10:333333:Calibri')

# ==========================================================================
# Sheet: 2-Area Variants
# ==========================================================================
print("\n--- 2-Area Variants ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Area Variants"')

# --------------------------------------------------------------------------
# Chart 1: Stacked area with plotFill and rounded corners
#
# officecli add charts-area.xlsx "/2-Area Variants" --type chart \
#   --prop chartType=areaStacked \
#   --prop title="Cumulative Traffic Sources" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop plotFill=F5F5F5 \
#   --prop roundedCorners=true \
#   --prop legend=bottom
#
# Features: chartType=areaStacked, plotFill (solid), roundedCorners
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Area Variants" --type chart'
    f' --prop chartType=areaStacked'
    f' --prop title="Cumulative Traffic Sources"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop plotFill=F5F5F5'
    f' --prop roundedCorners=true'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: 100% stacked area with axis number format and axis line
#
# officecli add charts-area.xlsx "/2-Area Variants" --type chart \
#   --prop chartType=areaPercentStacked \
#   --prop title="Traffic Share by Channel" \
#   --prop dataRange=Sheet1!A1:E13 \
#   --prop colors=2E75B6,C55A11,548235,BF8F00 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop axisNumFmt=0% \
#   --prop axisLine=333333:1:solid \
#   --prop legend=bottom
#
# Features: chartType=areaPercentStacked, axisNumFmt, axisLine
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Area Variants" --type chart'
    f' --prop chartType=areaPercentStacked'
    f' --prop title="Traffic Share by Channel"'
    f' --prop dataRange=Sheet1!A1:E13'
    f' --prop colors=2E75B6,C55A11,548235,BF8F00'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop axisNumFmt=0%'
    f' --prop axisLine=333333:1:solid'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: 3D area with perspective rotation
#
# officecli add charts-area.xlsx "/2-Area Variants" --type chart \
#   --prop chartType=area3d \
#   --prop title="3D Regional Sales" \
#   --prop series1="East:120,135,148,162,155,178" \
#   --prop series2="West:95,108,115,128,142,155" \
#   --prop series3="Central:88,92,105,118,125,138" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop colors=4472C4,ED7D31,70AD47 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop view3d=20,25,15 \
#   --prop legend=right
#
# Features: chartType=area3d, view3d (rotX,rotY,perspective)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Area Variants" --type chart'
    f' --prop chartType=area3d'
    f' --prop title="3D Regional Sales"'
    f' --prop series1="East:120,135,148,162,155,178"'
    f' --prop series2="West:95,108,115,128,142,155"'
    f' --prop series3="Central:88,92,105,118,125,138"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop colors=4472C4,ED7D31,70AD47'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop view3d=20,25,15'
    f' --prop legend=right')

# --------------------------------------------------------------------------
# Chart 4: 3D stacked area
#
# officecli add charts-area.xlsx "/2-Area Variants" --type chart \
#   --prop chartType=area3d \
#   --prop title="3D Stacked Inventory" \
#   --prop series1="Warehouse A:500,480,520,550,530,560" \
#   --prop series2="Warehouse B:320,350,340,380,400,410" \
#   --prop series3="Warehouse C:180,200,210,230,250,240" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop colors=1F4E79,2E75B6,9DC3E6 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop view3d=15,20,20 \
#   --prop gridlines=D9D9D9:0.5:dot
#
# Features: area3d stacked appearance, multiple series, gridlines
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Area Variants" --type chart'
    f' --prop chartType=area3d'
    f' --prop title="3D Stacked Inventory"'
    f' --prop series1="Warehouse A:500,480,520,550,530,560"'
    f' --prop series2="Warehouse B:320,350,340,380,400,410"'
    f' --prop series3="Warehouse C:180,200,210,230,250,240"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop colors=1F4E79,2E75B6,9DC3E6'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop view3d=15,20,20'
    f' --prop gridlines=D9D9D9:0.5:dot')

# ==========================================================================
# Sheet: 3-Area Styling
# ==========================================================================
print("\n--- 3-Area Styling ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Area Styling"')

# --------------------------------------------------------------------------
# Chart 1: Title styling (font, size, color, bold, shadow)
#
# officecli add charts-area.xlsx "/3-Area Styling" --type chart \
#   --prop chartType=area \
#   --prop title="Styled Title Demo" \
#   --prop series1="Revenue:80,120,160,200,240,280" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop colors=4472C4 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop title.font=Georgia --prop title.size=16 \
#   --prop title.color=1F4E79 --prop title.bold=true \
#   --prop title.shadow=000000-3-315-2-30 \
#   --prop transparency=30
#
# Features: title.font, title.size, title.color, title.bold, title.shadow
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Area Styling" --type chart'
    f' --prop chartType=area'
    f' --prop title="Styled Title Demo"'
    f' --prop series1="Revenue:80,120,160,200,240,280"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop colors=4472C4'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop title.font=Georgia --prop title.size=16'
    f' --prop title.color=1F4E79 --prop title.bold=true'
    f' --prop title.shadow=000000-3-315-2-30'
    f' --prop transparency=30')

# --------------------------------------------------------------------------
# Chart 2: Series shadow, outline, and smooth curve
#
# officecli add charts-area.xlsx "/3-Area Styling" --type chart \
#   --prop chartType=area \
#   --prop title="Smooth Area with Effects" \
#   --prop series1="Signups:150,180,220,260,310,350" \
#   --prop series2="Trials:90,110,140,170,200,230" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop colors=4472C4,70AD47 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop smooth=true \
#   --prop series.shadow=000000-4-315-2-40 \
#   --prop series.outline=333333-1 \
#   --prop transparency=25
#
# Features: smooth, series.shadow (color-blur-angle-dist-opacity),
#   series.outline (color-width)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Area Styling" --type chart'
    f' --prop chartType=area'
    f' --prop title="Smooth Area with Effects"'
    f' --prop series1="Signups:150,180,220,260,310,350"'
    f' --prop series2="Trials:90,110,140,170,200,230"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop colors=4472C4,70AD47'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop smooth=true'
    f' --prop series.shadow=000000-4-315-2-40'
    f' --prop series.outline=333333-1'
    f' --prop transparency=25')

# --------------------------------------------------------------------------
# Chart 3: Axis font styling, gridlines, and minor gridlines
#
# officecli add charts-area.xlsx "/3-Area Styling" --type chart \
#   --prop chartType=area \
#   --prop title="Gridline Configuration" \
#   --prop dataRange=Sheet1!A1:C13 \
#   --prop colors=2E75B6,C55A11 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop axisfont=9:58626E:Arial \
#   --prop gridlines=D9D9D9:0.5:dot \
#   --prop minorGridlines=EEEEEE:0.3:dot \
#   --prop catTitle=Month --prop axisTitle=Visitors
#
# Features: axisfont (size:color:font), gridlines (color:width:dash),
#   minorGridlines
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Area Styling" --type chart'
    f' --prop chartType=area'
    f' --prop title="Gridline Configuration"'
    f' --prop dataRange=Sheet1!A1:C13'
    f' --prop colors=2E75B6,C55A11'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop axisfont=9:58626E:Arial'
    f' --prop gridlines=D9D9D9:0.5:dot'
    f' --prop minorGridlines=EEEEEE:0.3:dot'
    f' --prop catTitle=Month --prop axisTitle=Visitors')

# --------------------------------------------------------------------------
# Chart 4: Chart fill, plot fill gradient, chart/plot area borders
#
# officecli add charts-area.xlsx "/3-Area Styling" --type chart \
#   --prop chartType=area \
#   --prop title="Fills and Borders" \
#   --prop series1="Sales:200,240,280,320,360,400" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop colors=4472C4 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop chartFill=FAFAFA \
#   --prop plotFill=E8F0FE-D6E4F0:90 \
#   --prop chartArea.border=D0D0D0:1:solid \
#   --prop plotArea.border=E0E0E0:0.5:dot \
#   --prop roundedCorners=true
#
# Features: chartFill, plotFill (gradient from-to:angle),
#   chartArea.border, plotArea.border, roundedCorners
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Area Styling" --type chart'
    f' --prop chartType=area'
    f' --prop title="Fills and Borders"'
    f' --prop series1="Sales:200,240,280,320,360,400"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop colors=4472C4'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop chartFill=FAFAFA'
    f' --prop "plotFill=E8F0FE-D6E4F0:90"'
    f' --prop chartArea.border=D0D0D0:1:solid'
    f' --prop plotArea.border=E0E0E0:0.5:dot'
    f' --prop roundedCorners=true')

# ==========================================================================
# Sheet: 4-Labels & Legend
# ==========================================================================
print("\n--- 4-Labels & Legend ---")
cli(f'add "{FILE}" / --type sheet --prop name="4-Labels & Legend"')

# --------------------------------------------------------------------------
# Chart 1: Data labels with position, font, and number format
#
# officecli add charts-area.xlsx "/4-Labels & Legend" --type chart \
#   --prop chartType=area \
#   --prop title="Labeled Area Chart" \
#   --prop series1="Users:3200,3800,4500,5100,5800,6400" \
#   --prop categories=Jul,Aug,Sep,Oct,Nov,Dec \
#   --prop colors=4472C4 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelPos=top \
#   --prop labelFont=9:333333:true \
#   --prop dataLabels.numFmt=#,##0
#
# Features: dataLabels, labelPos (top), labelFont (size:color:bold),
#   dataLabels.numFmt
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Labels & Legend" --type chart'
    f' --prop chartType=area'
    f' --prop title="Labeled Area Chart"'
    f' --prop series1="Users:3200,3800,4500,5100,5800,6400"'
    f' --prop categories=Jul,Aug,Sep,Oct,Nov,Dec'
    f' --prop colors=4472C4'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop dataLabels=true --prop labelPos=top'
    f' --prop labelFont=9:333333:true'
    f' --prop dataLabels.numFmt=#,##0')

# --------------------------------------------------------------------------
# Chart 2: Individual label deletion and per-point colors
#
# officecli add charts-area.xlsx "/4-Labels & Legend" --type chart \
#   --prop chartType=area \
#   --prop title="Highlighted Peak Month" \
#   --prop series1="Revenue:180,210,250,310,280,260" \
#   --prop categories=Jul,Aug,Sep,Oct,Nov,Dec \
#   --prop colors=2E75B6 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop dataLabels=true \
#   --prop dataLabel1.delete=true --prop dataLabel2.delete=true \
#   --prop dataLabel5.delete=true --prop dataLabel6.delete=true \
#   --prop point4.color=C00000 \
#   --prop transparency=30
#
# Features: dataLabel{N}.delete, point{N}.color
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Labels & Legend" --type chart'
    f' --prop chartType=area'
    f' --prop title="Highlighted Peak Month"'
    f' --prop series1="Revenue:180,210,250,310,280,260"'
    f' --prop categories=Jul,Aug,Sep,Oct,Nov,Dec'
    f' --prop colors=2E75B6'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop dataLabels=true'
    f' --prop dataLabel1.delete=true --prop dataLabel2.delete=true'
    f' --prop dataLabel5.delete=true --prop dataLabel6.delete=true'
    f' --prop point4.color=C00000'
    f' --prop transparency=30')

# --------------------------------------------------------------------------
# Chart 3: Legend positioning with overlay and font styling
#
# officecli add charts-area.xlsx "/4-Labels & Legend" --type chart \
#   --prop chartType=area \
#   --prop title="Legend Overlay Demo" \
#   --prop series1="Desktop:4200,4800,5100,5600" \
#   --prop series2="Mobile:3100,3500,3800,4200" \
#   --prop series3="Tablet:1200,1400,1500,1700" \
#   --prop categories=Q1,Q2,Q3,Q4 \
#   --prop colors=4472C4,ED7D31,70AD47 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop legend=right --prop legendfont=10:1F4E79:Calibri \
#   --prop legend.overlay=true \
#   --prop transparency=35
#
# Features: legend=right, legendfont, legend.overlay
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Labels & Legend" --type chart'
    f' --prop chartType=area'
    f' --prop title="Legend Overlay Demo"'
    f' --prop series1="Desktop:4200,4800,5100,5600"'
    f' --prop series2="Mobile:3100,3500,3800,4200"'
    f' --prop series3="Tablet:1200,1400,1500,1700"'
    f' --prop categories=Q1,Q2,Q3,Q4'
    f' --prop colors=4472C4,ED7D31,70AD47'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop legend=right --prop legendfont=10:1F4E79:Calibri'
    f' --prop legend.overlay=true'
    f' --prop transparency=35')

# --------------------------------------------------------------------------
# Chart 4: Manual layout — plotArea positioning
#
# officecli add charts-area.xlsx "/4-Labels & Legend" --type chart \
#   --prop chartType=area \
#   --prop title="Manual Layout" \
#   --prop series1="Growth:100,130,170,220,280,350" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop colors=70AD47 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop plotArea.x=0.12 --prop plotArea.y=0.18 \
#   --prop plotArea.w=0.82 --prop plotArea.h=0.55 \
#   --prop title.x=0.25 --prop title.y=0.02 \
#   --prop legend.x=0.15 --prop legend.y=0.82 \
#   --prop legend.w=0.7 --prop legend.h=0.12
#
# Features: plotArea.x/y/w/h, title.x/y, legend.x/y/w/h (manual layout)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Labels & Legend" --type chart'
    f' --prop chartType=area'
    f' --prop title="Manual Layout"'
    f' --prop series1="Growth:100,130,170,220,280,350"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop colors=70AD47'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop plotArea.x=0.12 --prop plotArea.y=0.18'
    f' --prop plotArea.w=0.82 --prop plotArea.h=0.55'
    f' --prop title.x=0.25 --prop title.y=0.02'
    f' --prop legend.x=0.15 --prop legend.y=0.82'
    f' --prop legend.w=0.7 --prop legend.h=0.12')

# ==========================================================================
# Sheet: 5-Advanced
# ==========================================================================
print("\n--- 5-Advanced ---")
cli(f'add "{FILE}" / --type sheet --prop name="5-Advanced"')

# --------------------------------------------------------------------------
# Chart 1: Secondary axis (dual scale)
#
# officecli add charts-area.xlsx "/5-Advanced" --type chart \
#   --prop chartType=area \
#   --prop title="Revenue vs Conversion Rate" \
#   --prop series1="Revenue:120,180,250,310,280,340" \
#   --prop series2="Conv %:2.1,2.8,3.2,3.9,3.5,4.1" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop colors=4472C4,C00000 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop secondaryAxis=2 \
#   --prop transparency=30
#
# Features: secondaryAxis (1-based series index on secondary Y axis)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Advanced" --type chart'
    f' --prop chartType=area'
    f' --prop title="Revenue vs Conversion Rate"'
    f' --prop series1="Revenue:120,180,250,310,280,340"'
    f' --prop series2="Conv %:2.1,2.8,3.2,3.9,3.5,4.1"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop colors=4472C4,C00000'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop secondaryAxis=2'
    f' --prop transparency=30')

# --------------------------------------------------------------------------
# Chart 2: Reference line
#
# officecli add charts-area.xlsx "/5-Advanced" --type chart \
#   --prop chartType=area \
#   --prop title="Sales vs Target" \
#   --prop series1="Sales:85,92,108,115,98,120" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop colors=4472C4 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop referenceLine=100:FF0000:1.5:dash \
#   --prop transparency=25 \
#   --prop areafill=4472C4-BDD7EE:90
#
# Features: referenceLine (value:color:width:dash)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Advanced" --type chart'
    f' --prop chartType=area'
    f' --prop title="Sales vs Target"'
    f' --prop series1="Sales:85,92,108,115,98,120"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop colors=4472C4'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop referenceLine=100:FF0000:1.5:dash'
    f' --prop transparency=25'
    f' --prop areafill=4472C4-BDD7EE:90')

# --------------------------------------------------------------------------
# Chart 3: Axis min/max, major unit, log scale, display units
#
# officecli add charts-area.xlsx "/5-Advanced" --type chart \
#   --prop chartType=area \
#   --prop title="Axis Scaling Demo" \
#   --prop series1="Visits:3200,3800,4500,5100,5800,6400" \
#   --prop categories=Jul,Aug,Sep,Oct,Nov,Dec \
#   --prop colors=2E75B6 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop axisMin=3000 --prop axisMax=7000 \
#   --prop majorUnit=500 \
#   --prop dispUnits=thousands \
#   --prop axisTitle=Visitors (K) \
#   --prop transparency=30
#
# Features: axisMin, axisMax, majorUnit, dispUnits (thousands/millions)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Advanced" --type chart'
    f' --prop chartType=area'
    f' --prop title="Axis Scaling Demo"'
    f' --prop series1="Visits:3200,3800,4500,5100,5800,6400"'
    f' --prop categories=Jul,Aug,Sep,Oct,Nov,Dec'
    f' --prop colors=2E75B6'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop axisMin=3000 --prop axisMax=7000'
    f' --prop majorUnit=500'
    f' --prop dispUnits=thousands'
    f' --prop "axisTitle=Visitors (K)"'
    f' --prop transparency=30')

# --------------------------------------------------------------------------
# Chart 4: Color rule, title glow, series shadow
#
# officecli add charts-area.xlsx "/5-Advanced" --type chart \
#   --prop chartType=area \
#   --prop title="Performance Threshold" \
#   --prop series1="Score:45,62,38,71,55,80" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop colorRule=50:C00000:70AD47 \
#   --prop referenceLine=50:888888:1:solid \
#   --prop title.glow=4472C4-8-60 \
#   --prop series.shadow=000000-3-315-1-30 \
#   --prop transparency=20
#
# Features: colorRule (threshold:belowColor:aboveColor), title.glow
#   (color-radius-opacity), series.shadow
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Advanced" --type chart'
    f' --prop chartType=area'
    f' --prop title="Performance Threshold"'
    f' --prop series1="Score:45,62,38,71,55,80"'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop colorRule=50:C00000:70AD47'
    f' --prop referenceLine=50:888888:1:solid'
    f' --prop title.glow=4472C4-8-60'
    f' --prop series.shadow=000000-3-315-1-30'
    f' --prop transparency=20')

print(f"\nDone! Generated: {FILE}")
print("  6 sheets (Sheet1 data + 5 chart sheets, 20 charts total)")
