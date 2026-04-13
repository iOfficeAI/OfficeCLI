#!/usr/bin/env python3
"""
Bubble Charts Showcase — bubble scale, size representation, and styling.

Generates: charts-bubble.xlsx

Usage:
  python3 charts-bubble.py
"""

import subprocess, sys, os, json, atexit

FILE = "charts-bubble.xlsx"

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
# Sheet: 1-Bubble Fundamentals
# ==========================================================================
print("\n--- 1-Bubble Fundamentals ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Bubble Fundamentals"')

# --------------------------------------------------------------------------
# Chart 1: Basic bubble chart with 2 series
#
# officecli add charts-bubble.xlsx "/1-Bubble Fundamentals" --type chart \
#   --prop chartType=bubble \
#   --prop title="Market Analysis" \
#   --prop series1="Enterprise:50,12,80;120,8,45;200,15,60" \
#   --prop series2="Consumer:30,25,50;80,18,35;150,22,70" \
#   --prop colors=4472C4,ED7D31 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop catTitle=Market Size --prop axisTitle=Growth Rate \
#   --prop legend=bottom
#
# Features: chartType=bubble, X;Y;Size triplets, catTitle, axisTitle
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Bubble Fundamentals" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Market Analysis"'
    f' --prop series1=Enterprise:80,45,60'
    f' --prop series2=Consumer:50,35,70'
    f' --prop colors=4472C4,ED7D31'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop "catTitle=Market Size" --prop "axisTitle=Growth Rate"'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: bubbleScale=100 with dataLabels
#
# officecli add charts-bubble.xlsx "/1-Bubble Fundamentals" --type chart \
#   --prop chartType=bubble \
#   --prop title="Product Portfolio" \
#   --prop series1="Products:20,30,90;60,20,50;100,10,70;140,25,40" \
#   --prop colors=2E75B6 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop bubbleScale=100 \
#   --prop dataLabels=true --prop labelPos=center \
#   --prop labelFont=9:FFFFFF:true \
#   --prop legend=bottom
#
# Features: bubbleScale=100, dataLabels with center positioning
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Bubble Fundamentals" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Product Portfolio"'
    f' --prop series1=Products:90,50,70,40'
    f' --prop colors=2E75B6'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop bubbleScale=100'
    f' --prop dataLabels=true --prop labelPos=center'
    f' --prop labelFont=9:FFFFFF:true'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: bubbleScale=50 vs bubbleScale=200 comparison (small scale)
#
# officecli add charts-bubble.xlsx "/1-Bubble Fundamentals" --type chart \
#   --prop chartType=bubble \
#   --prop title="Small Bubbles (Scale 50)" \
#   --prop series1="Tech:40,15,60;90,22,80;160,10,45" \
#   --prop series2="Finance:70,18,55;130,12,70;180,20,35" \
#   --prop colors=70AD47,FFC000 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop bubbleScale=50 \
#   --prop legend=bottom
#
# Features: bubbleScale=50 (smaller bubbles)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Bubble Fundamentals" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Small Bubbles (Scale 50)"'
    f' --prop series1=Tech:60,80,45'
    f' --prop series2=Finance:55,70,35'
    f' --prop colors=70AD47,FFC000'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop bubbleScale=50'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: sizeRepresents=width
#
# officecli add charts-bubble.xlsx "/1-Bubble Fundamentals" --type chart \
#   --prop chartType=bubble \
#   --prop title="Size by Width" \
#   --prop series1="Regions:35,28,70;85,15,40;140,20,55;190,30,85" \
#   --prop colors=5B9BD5 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop sizeRepresents=width \
#   --prop bubbleScale=100 \
#   --prop legend=bottom
#
# Features: sizeRepresents=width (bubble diameter proportional to value)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Bubble Fundamentals" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Size by Width"'
    f' --prop series1=Regions:70,40,55,85'
    f' --prop colors=5B9BD5'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop sizeRepresents=width'
    f' --prop bubbleScale=100'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 2-Bubble Styling
# ==========================================================================
print("\n--- 2-Bubble Styling ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Bubble Styling"')

# --------------------------------------------------------------------------
# Chart 1: Title styling, legend positioning
#
# officecli add charts-bubble.xlsx "/2-Bubble Styling" --type chart \
#   --prop chartType=bubble \
#   --prop title="Styled Bubble Chart" \
#   --prop series1="Segment A:45,20,65;100,15,50;160,25,80" \
#   --prop series2="Segment B:60,30,45;120,10,60;175,18,40" \
#   --prop colors=1F4E79,C55A11 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop title.font=Georgia --prop title.size=16 \
#   --prop title.color=1F4E79 --prop title.bold=true \
#   --prop legend=right --prop legendfont=10:333333:Calibri
#
# Features: title.font/size/color/bold, legend=right, legendfont
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bubble Styling" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Styled Bubble Chart"'
    f' --prop series1=SegmentA:65,50,80'
    f' --prop series2=SegmentB:45,60,40'
    f' --prop colors=1F4E79,C55A11'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop title.font=Georgia --prop title.size=16'
    f' --prop title.color=1F4E79 --prop title.bold=true'
    f' --prop legend=right --prop legendfont=10:333333:Calibri')

# --------------------------------------------------------------------------
# Chart 2: Series colors, transparency
#
# officecli add charts-bubble.xlsx "/2-Bubble Styling" --type chart \
#   --prop chartType=bubble \
#   --prop title="Transparent Overlapping Bubbles" \
#   --prop series1="Group X:30,25,75;70,30,60;110,15,90;150,22,50" \
#   --prop series2="Group Y:50,20,65;90,28,55;130,18,80;170,12,45" \
#   --prop colors=804472C4,80ED7D31 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop bubbleScale=120 \
#   --prop legend=bottom
#
# Features: ARGB colors with alpha (80=50% transparency)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bubble Styling" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Transparent Overlapping Bubbles"'
    f' --prop series1=GroupX:75,60,90,50'
    f' --prop series2=GroupY:65,55,80,45'
    f' --prop colors=804472C4,80ED7D31'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop bubbleScale=120'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: gridlines, axisfont, axisLine
#
# officecli add charts-bubble.xlsx "/2-Bubble Styling" --type chart \
#   --prop chartType=bubble \
#   --prop title="Grid & Axis Styling" \
#   --prop series1="Division 1:25,35,55;65,20,70;115,28,45" \
#   --prop series2="Division 2:40,15,60;80,25,40;130,30,75" \
#   --prop colors=2E75B6,548235 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop gridlines=D9D9D9:0.5 \
#   --prop axisfont=9:666666 \
#   --prop axisLine=333333-1 \
#   --prop legend=bottom
#
# Features: gridlines, axisfont, axisLine
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bubble Styling" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Grid & Axis Styling"'
    f' --prop series1=Div1:55,70,45'
    f' --prop series2=Div2:60,40,75'
    f' --prop colors=2E75B6,548235'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop gridlines=D9D9D9:0.5'
    f' --prop axisfont=9:666666'
    f' --prop axisLine=333333:1'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: plotFill, chartFill, series.shadow
#
# officecli add charts-bubble.xlsx "/2-Bubble Styling" --type chart \
#   --prop chartType=bubble \
#   --prop title="Shadow & Fill Effects" \
#   --prop series1="Portfolio:35,22,80;75,28,55;120,16,65;165,32,45" \
#   --prop colors=4472C4 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop plotFill=F0F4F8 --prop chartFill=FAFAFA \
#   --prop series.shadow=000000-4-315-2-30 \
#   --prop legend=bottom
#
# Features: plotFill, chartFill, series.shadow
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bubble Styling" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Shadow & Fill Effects"'
    f' --prop series1=Portfolio:80,55,65,45'
    f' --prop colors=4472C4'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop plotFill=F0F4F8 --prop chartFill=FAFAFA'
    f' --prop series.shadow=000000-4-315-2-30'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 3-Bubble Advanced
# ==========================================================================
print("\n--- 3-Bubble Advanced ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Bubble Advanced"')

# --------------------------------------------------------------------------
# Chart 1: secondaryAxis
#
# officecli add charts-bubble.xlsx "/3-Bubble Advanced" --type chart \
#   --prop chartType=bubble \
#   --prop title="Dual-Axis Bubble" \
#   --prop series1="Domestic:70,85,60,90" \
#   --prop series2="International:45,55,80,65" \
#   --prop categories=1,2,3,4 \
#   --prop colors=4472C4,ED7D31 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop secondaryAxis=2 \
#   --prop legend=bottom
#
# Features: secondaryAxis on bubble chart
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Bubble Advanced" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Dual-Axis Bubble"'
    f' --prop series1=Domestic:70,85,60,90'
    f' --prop series2=International:45,55,80,65'
    f' --prop categories=1,2,3,4'
    f' --prop colors=4472C4,ED7D31'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop secondaryAxis=2'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: referenceLine
#
# officecli add charts-bubble.xlsx "/3-Bubble Advanced" --type chart \
#   --prop chartType=bubble \
#   --prop title="Growth Threshold" \
#   --prop series1="Products:60,80,45,55" \
#   --prop categories=1,2,3,4 \
#   --prop colors=70AD47 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop referenceLine=50:C00000:Target \
#   --prop bubbleScale=80 \
#   --prop legend=bottom
#
# Features: referenceLine on bubble chart
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Bubble Advanced" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Growth Threshold"'
    f' --prop series1=Products:60,80,45,55'
    f' --prop categories=1,2,3,4'
    f' --prop colors=70AD47'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop "referenceLine=50:C00000:Target"'
    f' --prop bubbleScale=80'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: axisMin/Max, logBase
#
# officecli add charts-bubble.xlsx "/3-Bubble Advanced" --type chart \
#   --prop chartType=bubble \
#   --prop title="Log Scale Analysis" \
#   --prop series1="Markets:5,15,50,120" \
#   --prop categories=1,2,3,4 \
#   --prop colors=2E75B6 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop axisMin=1 --prop axisMax=200 \
#   --prop logBase=10 \
#   --prop bubbleScale=80 \
#   --prop legend=bottom
#
# Features: axisMin/Max, logBase=10 (logarithmic scale)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Bubble Advanced" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Log Scale Analysis"'
    f' --prop series1=Markets:5,15,50,120'
    f' --prop categories=1,2,3,4'
    f' --prop colors=2E75B6'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop axisMin=1 --prop axisMax=200'
    f' --prop logBase=10'
    f' --prop bubbleScale=80'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: chartArea.border, plotArea.border, trendline
#
# officecli add charts-bubble.xlsx "/3-Bubble Advanced" --type chart \
#   --prop chartType=bubble \
#   --prop title="Trend & Borders" \
#   --prop series1="Investments:20,55,95,140,180" \
#   --prop categories=1,2,3,4,5 \
#   --prop colors=4472C4 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop chartArea.border=333333:1.5 \
#   --prop plotArea.border=999999:0.75 \
#   --prop trendline=linear \
#   --prop legend=bottom
#
# Features: chartArea.border, plotArea.border, trendline=linear
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Bubble Advanced" --type chart'
    f' --prop chartType=bubble'
    f' --prop title="Trend & Borders"'
    f' --prop series1=Investments:20,55,95,140,180'
    f' --prop categories=1,2,3,4,5'
    f' --prop colors=4472C4'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop chartArea.border=333333:1.5'
    f' --prop plotArea.border=999999:0.75'
    f' --prop trendline=linear'
    f' --prop legend=bottom')

# Remove blank default Sheet1 (all data is inline)
cli(f'remove "{FILE}" /Sheet1')

print(f"\nDone! Generated: {FILE}")
print("  4 sheets (3 chart sheets, 12 charts total)")
