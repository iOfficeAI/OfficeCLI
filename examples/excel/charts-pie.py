#!/usr/bin/env python3
"""
Pie & Doughnut Charts Showcase — pie, pie3d, and doughnut with all variations.

Generates: charts-pie.xlsx

Usage:
  python3 charts-pie.py
"""

import subprocess, sys, os, json, atexit

FILE = "charts-pie.xlsx"

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
# Sheet: 1-Pie Charts
# ==========================================================================
print("\n--- 1-Pie Charts ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Pie Charts"')

# --------------------------------------------------------------------------
# Chart 1: Basic pie chart with inline data and custom colors
#
# officecli add charts-pie.xlsx "/1-Pie Charts" --type chart \
#   --prop chartType=pie \
#   --prop title="Market Share" \
#   --prop series1="Share:40,25,20,15" \
#   --prop categories=Product A,Product B,Product C,Product D \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelPos=outsideEnd
#
# Features: chartType=pie, inline series, categories, colors, dataLabels
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Pie Charts" --type chart'
    f' --prop chartType=pie'
    f' --prop title="Market Share"'
    f' --prop series1=Share:40,25,20,15'
    f' --prop categories=Product A,Product B,Product C,Product D'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop dataLabels=true --prop labelPos=outsideEnd')

# --------------------------------------------------------------------------
# Chart 2: Pie with exploded slice and per-point colors
#
# officecli add charts-pie.xlsx "/1-Pie Charts" --type chart \
#   --prop chartType=pie \
#   --prop title="Revenue by Region" \
#   --prop series1="Revenue:35,28,22,15" \
#   --prop categories=North,South,East,West \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop explosion=15 \
#   --prop point1.color=1F4E79 --prop point2.color=2E75B6 \
#   --prop point3.color=9DC3E6 --prop point4.color=BDD7EE \
#   --prop dataLabels=percent --prop labelPos=bestFit
#
# Features: explosion (slice separation %), point{N}.color, labelPos=bestFit,
#   dataLabels=percent
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Pie Charts" --type chart'
    f' --prop chartType=pie'
    f' --prop title="Revenue by Region"'
    f' --prop series1=Revenue:35,28,22,15'
    f' --prop categories=North,South,East,West'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop explosion=15'
    f' --prop point1.color=1F4E79 --prop point2.color=2E75B6'
    f' --prop point3.color=9DC3E6 --prop point4.color=BDD7EE'
    f' --prop dataLabels=true --prop labelPos=bestFit'
    f' --prop dataLabels=percent --prop labelPos=bestFit')

# --------------------------------------------------------------------------
# Chart 3: 3D pie with perspective and title styling
#
# officecli add charts-pie.xlsx "/1-Pie Charts" --type chart \
#   --prop chartType=pie3d \
#   --prop title="3D Category Split" \
#   --prop series1="Sales:45,30,25" \
#   --prop categories=Electronics,Clothing,Food \
#   --prop colors=2E75B6,70AD47,FFC000 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop view3d=30,0,0 \
#   --prop title.font=Georgia --prop title.size=16 \
#   --prop title.color=1F4E79 --prop title.bold=true \
#   --prop dataLabels=true --prop labelPos=center \
#   --prop labelFont=12:FFFFFF:true
#
# Features: pie3d, view3d on pie (tilt angle), title.font/size/color/bold,
#   labelFont (size:color:bold)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Pie Charts" --type chart'
    f' --prop chartType=pie3d'
    f' --prop title="3D Category Split"'
    f' --prop series1=Sales:45,30,25'
    f' --prop categories=Electronics,Clothing,Food'
    f' --prop colors=2E75B6,70AD47,FFC000'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop view3d=30,0,0'
    f' --prop title.font=Georgia --prop title.size=16'
    f' --prop title.color=1F4E79 --prop title.bold=true'
    f' --prop dataLabels=true --prop labelPos=center'
    f' --prop labelFont=12:FFFFFF:true')

# --------------------------------------------------------------------------
# Chart 4: Pie with gradient fills, leader lines, and legend positioning
#
# officecli add charts-pie.xlsx "/1-Pie Charts" --type chart \
#   --prop chartType=pie \
#   --prop title="Q4 Distribution" \
#   --prop series1="Q4:198,158,142,180" \
#   --prop categories=East,South,North,West \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90;70AD47-C5E0B4:90;FFC000-FFF2CC:90' \
#   --prop legend=right --prop legendfont=10:333333:Helvetica \
#   --prop dataLabels=true \
#   --prop dataLabels.showLeaderLines=true \
#   --prop chartFill=FAFAFA --prop roundedCorners=true
#
# Features: gradients (per-slice), legend=right, legendfont,
#   dataLabels.showLeaderLines, chartFill, roundedCorners
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Pie Charts" --type chart'
    f' --prop chartType=pie'
    f' --prop title="Q4 Distribution"'
    f' --prop series1=Q4:198,158,142,180'
    f' --prop categories=East,South,North,West'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop "gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90;70AD47-C5E0B4:90;FFC000-FFF2CC:90"'
    f' --prop legend=right --prop legendfont=10:333333:Helvetica'
    f' --prop dataLabels=true'
    f' --prop dataLabels.showLeaderLines=true'
    f' --prop chartFill=FAFAFA --prop roundedCorners=true')

# ==========================================================================
# Sheet: 2-Doughnut Charts
# ==========================================================================
print("\n--- 2-Doughnut Charts ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Doughnut Charts"')

# --------------------------------------------------------------------------
# Chart 1: Basic doughnut chart
#
# officecli add charts-pie.xlsx "/2-Doughnut Charts" --type chart \
#   --prop chartType=doughnut \
#   --prop title="Channel Mix" \
#   --prop series1="Channel:55,45" \
#   --prop categories=Online,Retail \
#   --prop colors=4472C4,ED7D31 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelPos=center \
#   --prop labelFont=14:FFFFFF:true
#
# Features: chartType=doughnut, center labels
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Doughnut Charts" --type chart'
    f' --prop chartType=doughnut'
    f' --prop title="Channel Mix"'
    f' --prop series1=Channel:55,45'
    f' --prop categories=Online,Retail'
    f' --prop colors=4472C4,ED7D31'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop dataLabels=true --prop labelPos=center'
    f' --prop labelFont=14:FFFFFF:true')

# --------------------------------------------------------------------------
# Chart 2: Multi-ring doughnut (multiple series)
#
# officecli add charts-pie.xlsx "/2-Doughnut Charts" --type chart \
#   --prop chartType=doughnut \
#   --prop title="Year-over-Year Comparison" \
#   --prop series1="2024:40,35,25" \
#   --prop series2="2025:45,30,25" \
#   --prop categories=Electronics,Clothing,Food \
#   --prop colors=4472C4,70AD47,FFC000 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop series.outline=FFFFFF-1 \
#   --prop legend=bottom
#
# Features: multi-ring doughnut (multiple series = concentric rings),
#   series.outline (white separator between slices)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Doughnut Charts" --type chart'
    f' --prop chartType=doughnut'
    f' --prop title="Year-over-Year Comparison"'
    f' --prop series1=2024:40,35,25'
    f' --prop series2=2025:45,30,25'
    f' --prop categories=Electronics,Clothing,Food'
    f' --prop colors=4472C4,70AD47,FFC000'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop series.outline=FFFFFF-1'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Styled doughnut with shadow and custom data labels
#
# officecli add charts-pie.xlsx "/2-Doughnut Charts" --type chart \
#   --prop chartType=doughnut \
#   --prop title="Priority Breakdown" \
#   --prop series1="Priority:50,30,20" \
#   --prop categories=High,Medium,Low \
#   --prop colors=C00000,FFC000,70AD47 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop series.shadow=000000-4-315-2-30 \
#   --prop dataLabels=true --prop labelPos=outsideEnd \
#   --prop dataLabels.numFmt=0"%" \
#   --prop title.shadow=000000-3-315-2-30 \
#   --prop plotFill=F5F5F5
#
# Features: series.shadow on doughnut, title.shadow, plotFill
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Doughnut Charts" --type chart'
    f' --prop chartType=doughnut'
    f' --prop title="Priority Breakdown"'
    f' --prop series1=Priority:50,30,20'
    f' --prop categories=High,Medium,Low'
    f' --prop colors=C00000,FFC000,70AD47'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop series.shadow=000000-4-315-2-30'
    f' --prop dataLabels=true --prop labelPos=outsideEnd'
    f' --prop dataLabels.numFmt=0"%"'
    f' --prop title.shadow=000000-3-315-2-30'
    f' --prop plotFill=F5F5F5')

# --------------------------------------------------------------------------
# Chart 4: Doughnut with per-slice gradient and explosion
#
# officecli add charts-pie.xlsx "/2-Doughnut Charts" --type chart \
#   --prop chartType=doughnut \
#   --prop title="Product Revenue" \
#   --prop series1="Revenue:35,25,20,12,8" \
#   --prop categories=Laptop,Phone,Tablet,Jacket,Coffee \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop explosion=8 \
#   --prop 'gradients=1F4E79-5B9BD5:90;C55A11-F4B183:90;548235-A9D18E:90;7F6000-FFD966:90;843C0B-DDA15E:90' \
#   --prop legend=right \
#   --prop dataLabels=true --prop labelPos=bestFit
#
# Features: explosion on doughnut, 5-slice gradients
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Doughnut Charts" --type chart'
    f' --prop chartType=doughnut'
    f' --prop title="Product Revenue"'
    f' --prop series1=Revenue:35,25,20,12,8'
    f' --prop categories=Laptop,Phone,Tablet,Jacket,Coffee'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop explosion=8'
    f' --prop "gradients=1F4E79-5B9BD5:90;C55A11-F4B183:90;548235-A9D18E:90;7F6000-FFD966:90;843C0B-DDA15E:90"'
    f' --prop legend=right'
    f' --prop dataLabels=true --prop labelPos=bestFit')

# Remove blank default Sheet1 (all data is inline)
cli(f'remove "{FILE}" /Sheet1')

print(f"\nDone! Generated: {FILE}")
print("  3 sheets (2 chart sheets, 8 charts total)")
