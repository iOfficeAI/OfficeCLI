#!/usr/bin/env python3
"""
Radar Charts Showcase — radar with standard, filled, and marker styles.

Generates: charts-radar.xlsx

Usage:
  python3 charts-radar.py
"""

import subprocess, sys, os, json, atexit

FILE = "charts-radar.xlsx"

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
# Sheet: 1-Radar Fundamentals
# ==========================================================================
print("\n--- 1-Radar Fundamentals ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Radar Fundamentals"')

# --------------------------------------------------------------------------
# Chart 1: Basic radar (standard style) with 3 series
#
# officecli add charts-radar.xlsx "/1-Radar Fundamentals" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=standard \
#   --prop title="Athlete Comparison" \
#   --prop series1="Alice:85,70,90,60,75" \
#   --prop series2="Bob:65,90,70,80,85" \
#   --prop series3="Carol:75,80,80,70,65" \
#   --prop categories=Speed,Strength,Stamina,Agility,Accuracy \
#   --prop colors=4472C4,ED7D31,70AD47 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop legend=bottom
#
# Features: chartType=radar, radarStyle=standard, 3 series, categories as spokes
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Radar Fundamentals" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=standard'
    f' --prop title="Athlete Comparison"'
    f' --prop series1=Alice:85,70,90,60,75'
    f' --prop series2=Bob:65,90,70,80,85'
    f' --prop series3=Carol:75,80,80,70,65'
    f' --prop categories=Speed,Strength,Stamina,Agility,Accuracy'
    f' --prop colors=4472C4,ED7D31,70AD47'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: Radar with markers (marker style)
#
# officecli add charts-radar.xlsx "/1-Radar Fundamentals" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=marker \
#   --prop title="Product Ratings" \
#   --prop series1="Product A:9,7,8,6,8" \
#   --prop series2="Product B:6,9,7,8,5" \
#   --prop categories=Quality,Price,Design,Support,Delivery \
#   --prop colors=2E75B6,C00000 \
#   --prop marker=circle:6:2E75B6 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop legend=bottom \
#   --prop dataLabels=true
#
# Features: radarStyle=marker, marker=circle:6:color, dataLabels
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Radar Fundamentals" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=marker'
    f' --prop title="Product Ratings"'
    f' --prop series1="Product A:9,7,8,6,8"'
    f' --prop series2="Product B:6,9,7,8,5"'
    f' --prop categories=Quality,Price,Design,Support,Delivery'
    f' --prop colors=2E75B6,C00000'
    f' --prop marker=circle:6:2E75B6'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop legend=bottom'
    f' --prop dataLabels=true')

# --------------------------------------------------------------------------
# Chart 3: Filled radar with transparency
#
# officecli add charts-radar.xlsx "/1-Radar Fundamentals" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=filled \
#   --prop title="Skills Assessment" \
#   --prop series1="Junior:50,40,60,70,55" \
#   --prop series2="Senior:85,80,75,90,80" \
#   --prop categories=Coding,Design,Testing,Communication,Leadership \
#   --prop colors=4472C4,70AD47 \
#   --prop transparency=40 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop legend=bottom
#
# Features: radarStyle=filled, transparency=40 (semi-transparent fill)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Radar Fundamentals" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=filled'
    f' --prop title="Skills Assessment"'
    f' --prop series1=Junior:50,40,60,70,55'
    f' --prop series2=Senior:85,80,75,90,80'
    f' --prop categories=Coding,Design,Testing,Communication,Leadership'
    f' --prop colors=4472C4,70AD47'
    f' --prop transparency=40'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: Filled radar with per-series colors and white outline
#
# officecli add charts-radar.xlsx "/1-Radar Fundamentals" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=filled \
#   --prop title="Department Scores" \
#   --prop series1="Engineering:90,75,60,85,70" \
#   --prop series2="Marketing:60,85,80,70,90" \
#   --prop series3="Sales:70,80,75,65,85" \
#   --prop categories=Innovation,Teamwork,Efficiency,Quality,Growth \
#   --prop colors=4472C4,ED7D31,70AD47 \
#   --prop series.outline=FFFFFF-0.5 \
#   --prop transparency=35 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop legend=bottom
#
# Features: filled radar, series.outline (white border between areas),
#   3 overlapping series with transparency
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Radar Fundamentals" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=filled'
    f' --prop title="Department Scores"'
    f' --prop series1=Engineering:90,75,60,85,70'
    f' --prop series2=Marketing:60,85,80,70,90'
    f' --prop series3=Sales:70,80,75,65,85'
    f' --prop categories=Innovation,Teamwork,Efficiency,Quality,Growth'
    f' --prop colors=4472C4,ED7D31,70AD47'
    f' --prop series.outline=FFFFFF-0.5'
    f' --prop transparency=35'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 2-Radar Styling
# ==========================================================================
print("\n--- 2-Radar Styling ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Radar Styling"')

# --------------------------------------------------------------------------
# Chart 1: Title styling (font, size, color, bold, shadow)
#
# officecli add charts-radar.xlsx "/2-Radar Styling" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=marker \
#   --prop title="Styled Title Demo" \
#   --prop series1="Team A:80,65,90,70,85" \
#   --prop categories=Attack,Defense,Speed,Skill,Stamina \
#   --prop colors=2E75B6 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop title.font=Georgia --prop title.size=18 \
#   --prop title.color=1F4E79 --prop title.bold=true \
#   --prop title.shadow=000000-3-315-2-30
#
# Features: title.font, title.size, title.color, title.bold, title.shadow
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Radar Styling" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=marker'
    f' --prop title="Styled Title Demo"'
    f' --prop series1="Team A:80,65,90,70,85"'
    f' --prop categories=Attack,Defense,Speed,Skill,Stamina'
    f' --prop colors=2E75B6'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop title.font=Georgia --prop title.size=18'
    f' --prop title.color=1F4E79 --prop title.bold=true'
    f' --prop title.shadow=000000-3-315-2-30')

# --------------------------------------------------------------------------
# Chart 2: Series shadow effects
#
# officecli add charts-radar.xlsx "/2-Radar Styling" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=filled \
#   --prop title="Shadow Effects" \
#   --prop series1="Region A:75,80,65,90,70" \
#   --prop series2="Region B:60,70,85,75,80" \
#   --prop categories=Revenue,Profit,Growth,Retention,Satisfaction \
#   --prop colors=4472C4,ED7D31 \
#   --prop series.shadow=000000-4-315-2-30 \
#   --prop transparency=30 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop legend=bottom
#
# Features: series.shadow on filled radar, transparency with shadow
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Radar Styling" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=filled'
    f' --prop title="Shadow Effects"'
    f' --prop series1="Region A:75,80,65,90,70"'
    f' --prop series2="Region B:60,70,85,75,80"'
    f' --prop categories=Revenue,Profit,Growth,Retention,Satisfaction'
    f' --prop colors=4472C4,ED7D31'
    f' --prop series.shadow=000000-4-315-2-30'
    f' --prop transparency=30'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Axis font and gridlines styling
#
# officecli add charts-radar.xlsx "/2-Radar Styling" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=marker \
#   --prop title="Axis & Gridlines" \
#   --prop series1="Actual:70,85,60,75,80" \
#   --prop series2="Target:80,80,80,80,80" \
#   --prop categories=KPI 1,KPI 2,KPI 3,KPI 4,KPI 5 \
#   --prop colors=4472C4,C00000 \
#   --prop axisfont=10:333333:Calibri \
#   --prop gridlines=D9D9D9:0.5 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop legend=bottom
#
# Features: axisfont (size:color:fontFamily), gridlines (color-width)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Radar Styling" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=marker'
    f' --prop title="Axis & Gridlines"'
    f' --prop series1=Actual:70,85,60,75,80'
    f' --prop series2=Target:80,80,80,80,80'
    f' --prop categories=KPI 1,KPI 2,KPI 3,KPI 4,KPI 5'
    f' --prop colors=4472C4,C00000'
    f' --prop axisfont=10:333333:Calibri'
    f' --prop gridlines=D9D9D9:0.5'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: Plot fill, chart fill, rounded corners, borders
#
# officecli add charts-radar.xlsx "/2-Radar Styling" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=filled \
#   --prop title="Chart Area Styling" \
#   --prop series1="Score:85,70,90,60,75" \
#   --prop categories=Speed,Power,Technique,Endurance,Flexibility \
#   --prop colors=4472C4 \
#   --prop transparency=25 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop plotFill=F5F5F5 --prop chartFill=FAFAFA \
#   --prop roundedCorners=true \
#   --prop chartArea.border=BFBFBF:0.5 \
#   --prop plotArea.border=D9D9D9:0.25
#
# Features: plotFill, chartFill, roundedCorners, chartArea.border,
#   plotArea.border
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Radar Styling" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=filled'
    f' --prop title="Chart Area Styling"'
    f' --prop series1=Score:85,70,90,60,75'
    f' --prop categories=Speed,Power,Technique,Endurance,Flexibility'
    f' --prop colors=4472C4'
    f' --prop transparency=25'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop plotFill=F5F5F5 --prop chartFill=FAFAFA'
    f' --prop roundedCorners=true'
    f' --prop chartArea.border=BFBFBF:0.5'
    f' --prop plotArea.border=D9D9D9:0.25')

# ==========================================================================
# Sheet: 3-Labels & Legend
# ==========================================================================
print("\n--- 3-Labels & Legend ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Labels & Legend"')

# --------------------------------------------------------------------------
# Chart 1: Data labels with font styling and position
#
# officecli add charts-radar.xlsx "/3-Labels & Legend" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=marker \
#   --prop title="Data Labels" \
#   --prop series1="Performance:88,72,95,67,81" \
#   --prop categories=Speed,Strength,Stamina,Agility,Accuracy \
#   --prop colors=2E75B6 \
#   --prop marker=circle:6:2E75B6 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelPos=outsideEnd \
#   --prop labelFont=9:333333:true
#
# Features: dataLabels=true, labelPos=outsideEnd, labelFont (size:color:bold)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Labels & Legend" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=marker'
    f' --prop title="Data Labels"'
    f' --prop series1=Performance:88,72,95,67,81'
    f' --prop categories=Speed,Strength,Stamina,Agility,Accuracy'
    f' --prop colors=2E75B6'
    f' --prop marker=circle:6:2E75B6'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop dataLabels=true --prop labelPos=outsideEnd'
    f' --prop labelFont=9:333333:true')

# --------------------------------------------------------------------------
# Chart 2: Legend positioning and styling with overlay
#
# officecli add charts-radar.xlsx "/3-Labels & Legend" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=standard \
#   --prop title="Legend Styles" \
#   --prop series1="Alpha:80,60,75,90,70" \
#   --prop series2="Beta:70,80,85,65,75" \
#   --prop series3="Gamma:65,75,70,80,85" \
#   --prop categories=Metric A,Metric B,Metric C,Metric D,Metric E \
#   --prop colors=4472C4,ED7D31,70AD47 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop legend=right \
#   --prop legendfont=10:1F4E79:Calibri \
#   --prop legend.overlay=true
#
# Features: legend=right, legendfont (size:color:fontFamily), legend.overlay
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Labels & Legend" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=standard'
    f' --prop title="Legend Styles"'
    f' --prop series1=Alpha:80,60,75,90,70'
    f' --prop series2=Beta:70,80,85,65,75'
    f' --prop series3=Gamma:65,75,70,80,85'
    f' --prop categories=Metric A,Metric B,Metric C,Metric D,Metric E'
    f' --prop colors=4472C4,ED7D31,70AD47'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop legend=right'
    f' --prop legendfont=10:1F4E79:Calibri'
    f' --prop legend.overlay=true')

# --------------------------------------------------------------------------
# Chart 3: Manual plot area layout
#
# officecli add charts-radar.xlsx "/3-Labels & Legend" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=filled \
#   --prop title="Custom Layout" \
#   --prop series1="Team:85,70,90,65,80" \
#   --prop categories=Vision,Execution,Culture,Agility,Impact \
#   --prop colors=4472C4 \
#   --prop transparency=30 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop plotArea.x=0.1 --prop plotArea.y=0.15 \
#   --prop plotArea.w=0.8 --prop plotArea.h=0.75
#
# Features: plotArea.x/y/w/h (fractional manual layout positioning)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Labels & Legend" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=filled'
    f' --prop title="Custom Layout"'
    f' --prop series1=Team:85,70,90,65,80'
    f' --prop categories=Vision,Execution,Culture,Agility,Impact'
    f' --prop colors=4472C4'
    f' --prop transparency=30'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop plotArea.x=0.1 --prop plotArea.y=0.15'
    f' --prop plotArea.w=0.8 --prop plotArea.h=0.75')

# --------------------------------------------------------------------------
# Chart 4: Multiple series (5+) comparison
#
# officecli add charts-radar.xlsx "/3-Labels & Legend" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=standard \
#   --prop title="Multi-Team Comparison" \
#   --prop series1="Dev:90,70,80,65,75" \
#   --prop series2="QA:60,85,70,80,90" \
#   --prop series3="Design:75,80,85,70,60" \
#   --prop series4="PM:80,65,75,90,70" \
#   --prop series5="DevOps:70,75,60,85,80" \
#   --prop categories=Speed,Quality,Innovation,Teamwork,Delivery \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000,7030A0 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop legend=bottom \
#   --prop legendfont=9:333333:Calibri
#
# Features: 5 series on one radar, distinguishing many overlapping lines
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Labels & Legend" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=standard'
    f' --prop title="Multi-Team Comparison"'
    f' --prop series1=Dev:90,70,80,65,75'
    f' --prop series2=QA:60,85,70,80,90'
    f' --prop series3=Design:75,80,85,70,60'
    f' --prop series4=PM:80,65,75,90,70'
    f' --prop series5=DevOps:70,75,60,85,80'
    f' --prop categories=Speed,Quality,Innovation,Teamwork,Delivery'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000,7030A0'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop legend=bottom'
    f' --prop legendfont=9:333333:Calibri')

# ==========================================================================
# Sheet: 4-Advanced
# ==========================================================================
print("\n--- 4-Advanced ---")
cli(f'add "{FILE}" / --type sheet --prop name="4-Advanced"')

# --------------------------------------------------------------------------
# Chart 1: Title glow and shadow effects
#
# officecli add charts-radar.xlsx "/4-Advanced" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=marker \
#   --prop title="Glow & Shadow Title" \
#   --prop series1="Score:75,85,65,90,80" \
#   --prop categories=Creativity,Logic,Memory,Focus,Speed \
#   --prop colors=2E75B6 \
#   --prop marker=diamond:7:2E75B6 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop title.font=Georgia --prop title.size=16 \
#   --prop title.bold=true --prop title.color=1F4E79 \
#   --prop title.glow=4472C4-8 \
#   --prop title.shadow=000000-3-315-2-30
#
# Features: title.glow (color-radius), title.shadow combined
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Advanced" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=marker'
    f' --prop title="Glow & Shadow Title"'
    f' --prop series1=Score:75,85,65,90,80'
    f' --prop categories=Creativity,Logic,Memory,Focus,Speed'
    f' --prop colors=2E75B6'
    f' --prop marker=diamond:7:2E75B6'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop title.font=Georgia --prop title.size=16'
    f' --prop title.bold=true --prop title.color=1F4E79'
    f' --prop title.glow=4472C4-8'
    f' --prop title.shadow=000000-3-315-2-30')

# --------------------------------------------------------------------------
# Chart 2: Radar with many spokes (8 categories)
#
# officecli add charts-radar.xlsx "/4-Advanced" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=filled \
#   --prop title="8-Spoke Assessment" \
#   --prop series1="Candidate:85,70,90,60,75,80,65,88" \
#   --prop series2="Benchmark:70,70,70,70,70,70,70,70" \
#   --prop categories=Technical,Communication,Leadership,Creativity,Analytical,Teamwork,Adaptability,Initiative \
#   --prop colors=4472C4,BFBFBF \
#   --prop transparency=35 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop legend=bottom \
#   --prop gridlines=D9D9D9:0.25
#
# Features: 8 categories (many spokes), benchmark overlay, gridlines
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Advanced" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=filled'
    f' --prop title="8-Spoke Assessment"'
    f' --prop series1=Candidate:85,70,90,60,75,80,65,88'
    f' --prop series2=Benchmark:70,70,70,70,70,70,70,70'
    f' --prop categories=Technical,Communication,Leadership,Creativity,Analytical,Teamwork,Adaptability,Initiative'
    f' --prop colors=4472C4,BFBFBF'
    f' --prop transparency=35'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop legend=bottom'
    f' --prop gridlines=D9D9D9:0.25')

# --------------------------------------------------------------------------
# Chart 3: Single-series radar with full styling
#
# officecli add charts-radar.xlsx "/4-Advanced" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=marker \
#   --prop title="Personal Profile" \
#   --prop series1="Self:92,78,85,65,88,70" \
#   --prop categories=Python,JavaScript,SQL,DevOps,Testing,Design \
#   --prop colors=7030A0 \
#   --prop marker=square:7:7030A0 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelFont=9:7030A0:true \
#   --prop title.font=Calibri --prop title.size=14 \
#   --prop title.color=7030A0 --prop title.bold=true \
#   --prop plotFill=F8F0FF --prop chartFill=FFFFFF \
#   --prop roundedCorners=true \
#   --prop chartArea.border=7030A0:0.5
#
# Features: single series with marker, full title/chart/plot styling,
#   themed color scheme (purple)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Advanced" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=marker'
    f' --prop title="Personal Profile"'
    f' --prop series1=Self:92,78,85,65,88,70'
    f' --prop categories=Python,JavaScript,SQL,DevOps,Testing,Design'
    f' --prop colors=7030A0'
    f' --prop marker=square:7:7030A0'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop dataLabels=true --prop labelFont=9:7030A0:true'
    f' --prop title.font=Calibri --prop title.size=14'
    f' --prop title.color=7030A0 --prop title.bold=true'
    f' --prop plotFill=F8F0FF --prop chartFill=FFFFFF'
    f' --prop roundedCorners=true'
    f' --prop chartArea.border=7030A0:0.5')

# --------------------------------------------------------------------------
# Chart 4: Two-series filled radar with low transparency for overlap
#
# officecli add charts-radar.xlsx "/4-Advanced" --type chart \
#   --prop chartType=radar \
#   --prop radarStyle=filled \
#   --prop title="Before vs After" \
#   --prop series1="Before:55,40,65,50,45" \
#   --prop series2="After:85,75,80,70,80" \
#   --prop categories=Revenue,Efficiency,Satisfaction,Innovation,Retention \
#   --prop colors=C00000,70AD47 \
#   --prop transparency=20 \
#   --prop series.outline=FFFFFF-0.75 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop legend=bottom \
#   --prop dataLabels=true --prop labelFont=9:333333:false \
#   --prop chartFill=FAFAFA --prop plotFill=F5F5F5
#
# Features: low transparency (20%) for visible overlap, before/after
#   comparison pattern, series.outline for separation
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Advanced" --type chart'
    f' --prop chartType=radar'
    f' --prop radarStyle=filled'
    f' --prop title="Before vs After"'
    f' --prop series1=Before:55,40,65,50,45'
    f' --prop series2=After:85,75,80,70,80'
    f' --prop categories=Revenue,Efficiency,Satisfaction,Innovation,Retention'
    f' --prop colors=C00000,70AD47'
    f' --prop transparency=20'
    f' --prop series.outline=FFFFFF-0.75'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop legend=bottom'
    f' --prop dataLabels=true --prop labelFont=9:333333:false'
    f' --prop chartFill=FAFAFA --prop plotFill=F5F5F5')

# Remove blank default Sheet1 (all data is inline)
cli(f'remove "{FILE}" /Sheet1')

print(f"\nDone! Generated: {FILE}")
print("  5 sheets (4 chart sheets, 16 charts total)")
