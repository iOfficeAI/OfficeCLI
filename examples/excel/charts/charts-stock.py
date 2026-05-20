#!/usr/bin/env python3
"""
Stock Charts Showcase — OHLC with hi-low lines, up-down bars, and styling.

Generates: charts-stock.xlsx

Usage:
  python3 charts-stock.py
"""

import subprocess, sys, os, json, atexit

FILE = "charts-stock.xlsx"

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
# Sheet: 1-Stock Fundamentals
# ==========================================================================
print("\n--- 1-Stock Fundamentals ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Stock Fundamentals"')

# --------------------------------------------------------------------------
# Chart 1: Basic OHLC stock chart
#
# officecli add charts-stock.xlsx "/1-Stock Fundamentals" --type chart \
#   --prop chartType=stock \
#   --prop title="ACME Corp Weekly OHLC" \
#   --prop series1="Open:142,145,148,150,147,152" \
#   --prop series2="High:148,151,155,156,153,158" \
#   --prop series3="Low:139,142,145,147,144,149" \
#   --prop series4="Close:145,148,150,147,152,155" \
#   --prop categories=Week 1,Week 2,Week 3,Week 4,Week 5,Week 6 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop catTitle=Week --prop axisTitle=Price ($) \
#   --prop legend=bottom
#
# Features: chartType=stock, 4 series (Open/High/Low/Close), catTitle, axisTitle
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Stock Fundamentals" --type chart'
    f' --prop chartType=stock'
    f' --prop title="ACME Corp Weekly OHLC"'
    f' --prop series1=Open:142,145,148,150,147,152'
    f' --prop series2=High:148,151,155,156,153,158'
    f' --prop series3=Low:139,142,145,147,144,149'
    f' --prop series4=Close:145,148,150,147,152,155'
    f' --prop "categories=Week 1,Week 2,Week 3,Week 4,Week 5,Week 6"'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop catTitle=Week --prop "axisTitle=Price ($)"'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: Stock with gridlines and axisfont
#
# officecli add charts-stock.xlsx "/1-Stock Fundamentals" --type chart \
#   --prop chartType=stock \
#   --prop title="Tech Sector Daily" \
#   --prop series1="Open:210,215,212,218,220" \
#   --prop series2="High:218,222,219,225,228" \
#   --prop series3="Low:207,211,208,214,216" \
#   --prop series4="Close:215,212,218,220,225" \
#   --prop categories=Mon,Tue,Wed,Thu,Fri \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop gridlines=D9D9D9:0.5 \
#   --prop axisfont=9:666666 \
#   --prop legend=bottom
#
# Features: gridlines, axisfont on stock chart
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Stock Fundamentals" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Tech Sector Daily"'
    f' --prop series1=Open:210,215,212,218,220'
    f' --prop series2=High:218,222,219,225,228'
    f' --prop series3=Low:207,211,208,214,216'
    f' --prop series4=Close:215,212,218,220,225'
    f' --prop categories=Mon,Tue,Wed,Thu,Fri'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop gridlines=D9D9D9:0.5'
    f' --prop axisfont=9:666666'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Stock with hiLowLines
#
# officecli add charts-stock.xlsx "/1-Stock Fundamentals" --type chart \
#   --prop chartType=stock \
#   --prop title="Energy Sector with Hi-Low Lines" \
#   --prop series1="Open:78,80,82,79,83,85" \
#   --prop series2="High:84,86,88,85,89,91" \
#   --prop series3="Low:75,77,79,76,80,82" \
#   --prop series4="Close:80,82,79,83,85,88" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop hiLowLines=true \
#   --prop legend=bottom
#
# Features: hiLowLines=true (vertical lines connecting high to low)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Stock Fundamentals" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Energy Sector with Hi-Low Lines"'
    f' --prop series1=Open:78,80,82,79,83,85'
    f' --prop series2=High:84,86,88,85,89,91'
    f' --prop series3=Low:75,77,79,76,80,82'
    f' --prop series4=Close:80,82,79,83,85,88'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop hiLowLines=true'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: Stock with updownbars
#
# officecli add charts-stock.xlsx "/1-Stock Fundamentals" --type chart \
#   --prop chartType=stock \
#   --prop title="Pharma Index with Up-Down Bars" \
#   --prop series1="Open:55,58,56,60,62,59" \
#   --prop series2="High:61,63,62,66,68,65" \
#   --prop series3="Low:52,55,53,57,59,56" \
#   --prop series4="Close:58,56,60,62,59,63" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop updownbars=100:70AD47:C00000 \
#   --prop legend=bottom
#
# Features: updownbars=gapWidth:upColor:downColor
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Stock Fundamentals" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Pharma Index with Up-Down Bars"'
    f' --prop series1=Open:55,58,56,60,62,59'
    f' --prop series2=High:61,63,62,66,68,65'
    f' --prop series3=Low:52,55,53,57,59,56'
    f' --prop series4=Close:58,56,60,62,59,63'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop updownbars=100:70AD47:C00000'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 2-Stock Styling
# ==========================================================================
print("\n--- 2-Stock Styling ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Stock Styling"')

# --------------------------------------------------------------------------
# Chart 1: Title styling, legend positioning
#
# officecli add charts-stock.xlsx "/2-Stock Styling" --type chart \
#   --prop chartType=stock \
#   --prop title="Styled Stock Chart" \
#   --prop series1="Open:165,170,168,172,175" \
#   --prop series2="High:175,178,176,180,183" \
#   --prop series3="Low:160,165,163,168,170" \
#   --prop series4="Close:170,168,172,175,180" \
#   --prop categories=Mon,Tue,Wed,Thu,Fri \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop title.font=Georgia --prop title.size=16 \
#   --prop title.color=1F4E79 --prop title.bold=true \
#   --prop legend=right --prop legendfont=10:333333:Calibri
#
# Features: title.font/size/color/bold, legend=right, legendfont
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Stock Styling" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Styled Stock Chart"'
    f' --prop series1=Open:165,170,168,172,175'
    f' --prop series2=High:175,178,176,180,183'
    f' --prop series3=Low:160,165,163,168,170'
    f' --prop series4=Close:170,168,172,175,180'
    f' --prop categories=Mon,Tue,Wed,Thu,Fri'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop title.font=Georgia --prop title.size=16'
    f' --prop title.color=1F4E79 --prop title.bold=true'
    f' --prop legend=right --prop legendfont=10:333333:Calibri')

# --------------------------------------------------------------------------
# Chart 2: Series effects, axisLine, catAxisLine
#
# officecli add charts-stock.xlsx "/2-Stock Styling" --type chart \
#   --prop chartType=stock \
#   --prop title="Axis Line Styling" \
#   --prop series1="Open:92,95,93,97,99" \
#   --prop series2="High:99,102,100,104,106" \
#   --prop series3="Low:88,91,89,93,95" \
#   --prop series4="Close:95,93,97,99,103" \
#   --prop categories=W1,W2,W3,W4,W5 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop hiLowLines=true \
#   --prop axisLine=333333:1.5 --prop catAxisLine=333333:1.5 \
#   --prop legend=bottom
#
# Features: axisLine, catAxisLine on stock chart
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Stock Styling" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Axis Line Styling"'
    f' --prop series1=Open:92,95,93,97,99'
    f' --prop series2=High:99,102,100,104,106'
    f' --prop series3=Low:88,91,89,93,95'
    f' --prop series4=Close:95,93,97,99,103'
    f' --prop categories=W1,W2,W3,W4,W5'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop hiLowLines=true'
    f' --prop axisLine=333333:1.5 --prop catAxisLine=333333:1.5'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: axisMin/Max, majorUnit
#
# officecli add charts-stock.xlsx "/2-Stock Styling" --type chart \
#   --prop chartType=stock \
#   --prop title="Custom Axis Range" \
#   --prop series1="Open:120,125,122,128,130" \
#   --prop series2="High:132,138,135,140,142" \
#   --prop series3="Low:115,120,118,124,126" \
#   --prop series4="Close:125,122,128,130,135" \
#   --prop categories=Day 1,Day 2,Day 3,Day 4,Day 5 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop axisMin=110 --prop axisMax=150 \
#   --prop majorUnit=10 \
#   --prop updownbars=100:70AD47:C00000 \
#   --prop legend=bottom
#
# Features: axisMin/Max, majorUnit
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Stock Styling" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Custom Axis Range"'
    f' --prop series1=Open:120,125,122,128,130'
    f' --prop series2=High:132,138,135,140,142'
    f' --prop series3=Low:115,120,118,124,126'
    f' --prop series4=Close:125,122,128,130,135'
    f' --prop "categories=Day 1,Day 2,Day 3,Day 4,Day 5"'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop axisMin=110 --prop axisMax=150'
    f' --prop majorUnit=10'
    f' --prop updownbars=100:70AD47:C00000'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: plotFill, chartFill, roundedCorners
#
# officecli add charts-stock.xlsx "/2-Stock Styling" --type chart \
#   --prop chartType=stock \
#   --prop title="Styled Chart Area" \
#   --prop series1="Open:48,50,52,49,53" \
#   --prop series2="High:55,57,59,56,60" \
#   --prop series3="Low:44,46,48,45,49" \
#   --prop series4="Close:50,52,49,53,56" \
#   --prop categories=Mon,Tue,Wed,Thu,Fri \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop plotFill=F0F4F8 --prop chartFill=FAFAFA \
#   --prop roundedCorners=true \
#   --prop hiLowLines=true \
#   --prop legend=bottom
#
# Features: plotFill, chartFill, roundedCorners
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Stock Styling" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Styled Chart Area"'
    f' --prop series1=Open:48,50,52,49,53'
    f' --prop series2=High:55,57,59,56,60'
    f' --prop series3=Low:44,46,48,45,49'
    f' --prop series4=Close:50,52,49,53,56'
    f' --prop categories=Mon,Tue,Wed,Thu,Fri'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop plotFill=F0F4F8 --prop chartFill=FAFAFA'
    f' --prop roundedCorners=true'
    f' --prop hiLowLines=true'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 3-Stock Advanced
# ==========================================================================
print("\n--- 3-Stock Advanced ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Stock Advanced"')

# --------------------------------------------------------------------------
# Chart 1: dataLabels, labelFont
#
# officecli add charts-stock.xlsx "/3-Stock Advanced" --type chart \
#   --prop chartType=stock \
#   --prop title="Stock with Data Labels" \
#   --prop series1="Open:185,190,188,192,195" \
#   --prop series2="High:195,198,196,200,203" \
#   --prop series3="Low:180,185,183,188,190" \
#   --prop series4="Close:190,188,192,195,200" \
#   --prop categories=W1,W2,W3,W4,W5 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelPos=top \
#   --prop labelFont=8:666666:false \
#   --prop legend=bottom
#
# Features: dataLabels, labelPos, labelFont on stock
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Stock Advanced" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Stock with Data Labels"'
    f' --prop series1=Open:185,190,188,192,195'
    f' --prop series2=High:195,198,196,200,203'
    f' --prop series3=Low:180,185,183,188,190'
    f' --prop series4=Close:190,188,192,195,200'
    f' --prop categories=W1,W2,W3,W4,W5'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop dataLabels=true --prop labelPos=top'
    f' --prop labelFont=8:666666:false'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: referenceLine (support/resistance)
#
# officecli add charts-stock.xlsx "/3-Stock Advanced" --type chart \
#   --prop chartType=stock \
#   --prop title="Support & Resistance" \
#   --prop series1="Open:105,108,106,110,112,109" \
#   --prop series2="High:112,115,113,117,119,116" \
#   --prop series3="Low:101,104,102,106,108,105" \
#   --prop series4="Close:108,106,110,112,109,113" \
#   --prop categories=Jan,Feb,Mar,Apr,May,Jun \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop referenceLine=115:C00000:Resistance \
#   --prop hiLowLines=true \
#   --prop legend=bottom
#
# Features: referenceLine as support/resistance level
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Stock Advanced" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Support & Resistance"'
    f' --prop series1=Open:105,108,106,110,112,109'
    f' --prop series2=High:112,115,113,117,119,116'
    f' --prop series3=Low:101,104,102,106,108,105'
    f' --prop series4=Close:108,106,110,112,109,113'
    f' --prop categories=Jan,Feb,Mar,Apr,May,Jun'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop referenceLine=115:C00000:Resistance'
    f' --prop hiLowLines=true'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: chartArea.border, plotArea.border
#
# officecli add charts-stock.xlsx "/3-Stock Advanced" --type chart \
#   --prop chartType=stock \
#   --prop title="Bordered Stock Chart" \
#   --prop series1="Open:72,75,73,77,79" \
#   --prop series2="High:79,82,80,84,86" \
#   --prop series3="Low:68,71,69,73,75" \
#   --prop series4="Close:75,73,77,79,83" \
#   --prop categories=Mon,Tue,Wed,Thu,Fri \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop chartArea.border=333333:1.5 \
#   --prop plotArea.border=999999:0.75 \
#   --prop updownbars=100:70AD47:C00000 \
#   --prop legend=bottom
#
# Features: chartArea.border, plotArea.border
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Stock Advanced" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Bordered Stock Chart"'
    f' --prop series1=Open:72,75,73,77,79'
    f' --prop series2=High:79,82,80,84,86'
    f' --prop series3=Low:68,71,69,73,75'
    f' --prop series4=Close:75,73,77,79,83'
    f' --prop categories=Mon,Tue,Wed,Thu,Fri'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop chartArea.border=333333:1.5'
    f' --prop plotArea.border=999999:0.75'
    f' --prop updownbars=100:70AD47:C00000'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: dispUnits, axisNumFmt
#
# officecli add charts-stock.xlsx "/3-Stock Advanced" --type chart \
#   --prop chartType=stock \
#   --prop title="Large Cap Stock" \
#   --prop series1="Open:2850,2900,2880,2920,2950" \
#   --prop series2="High:2950,2980,2960,3000,3020" \
#   --prop series3="Low:2800,2850,2830,2870,2900" \
#   --prop series4="Close:2900,2880,2920,2950,2990" \
#   --prop categories=Q1,Q2,Q3,Q4,Q5 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop axisNumFmt=$#,##0 \
#   --prop hiLowLines=true \
#   --prop legend=bottom
#
# Features: axisNumFmt (dollar format)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Stock Advanced" --type chart'
    f' --prop chartType=stock'
    f' --prop title="Large Cap Stock"'
    f' --prop series1=Open:2850,2900,2880,2920,2950'
    f' --prop series2=High:2950,2980,2960,3000,3020'
    f' --prop series3=Low:2800,2850,2830,2870,2900'
    f' --prop series4=Close:2900,2880,2920,2950,2990'
    f' --prop categories=Q1,Q2,Q3,Q4,Q5'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop "axisNumFmt=$#,##0"'
    f' --prop hiLowLines=true'
    f' --prop legend=bottom')

# Remove blank default Sheet1 (all data is inline)
cli(f'remove "{FILE}" /Sheet1')

print(f"\nDone! Generated: {FILE}")
print("  4 sheets (3 chart sheets, 12 charts total)")
