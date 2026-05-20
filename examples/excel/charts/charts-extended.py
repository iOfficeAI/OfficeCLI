#!/usr/bin/env python3
"""
Extended Chart Types Showcase — full feature coverage for waterfall, funnel,
treemap, sunburst, histogram, boxWhisker (cx:chart family).

Covers every extended-chart-specific property plus representative generic
cx styling knobs (title.glow, chartFill gradient, legendfont, dataLabels...).

Generates: charts-extended.xlsx

Usage:
  python3 charts-extended.py
"""

import subprocess, sys, os, atexit

FILE = "charts-extended.xlsx"

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
# Sheet 1: Waterfall & Funnel
# ==========================================================================
print("\n--- 1-Waterfall & Funnel ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Waterfall & Funnel"')

# --------------------------------------------------------------------------
# Chart 1: Waterfall — increase/decrease/total colors + data labels + title glow
#
# officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
#   --prop chartType=waterfall \
#   --prop title="Cash Flow Bridge" \
#   --prop data="Start:1000,Revenue:500,Costs:-300,Tax:-100,Net:1100" \
#   --prop increaseColor=70AD47 \
#   --prop decreaseColor=FF0000 \
#   --prop totalColor=4472C4 \
#   --prop dataLabels=true \
#   --prop title.glow="00D2FF-6-60" \
#   --prop x=0 --prop y=0 --prop width=13 --prop height=18
#
# Features: chartType=waterfall, increaseColor, decreaseColor, totalColor,
#   dataLabels, title.glow
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Waterfall & Funnel" --type chart'
    f' --prop chartType=waterfall'
    f' --prop title="Cash Flow Bridge"'
    f' --prop data=Start:1000,Revenue:500,Costs:-300,Tax:-100,Net:1100'
    f' --prop increaseColor=70AD47'
    f' --prop decreaseColor=FF0000'
    f' --prop totalColor=4472C4'
    f' --prop dataLabels=true'
    f' --prop title.glow=00D2FF-6-60'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 2: Waterfall — chart-area gradient fill + legend + custom label font
#
# officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
#   --prop chartType=waterfall \
#   --prop title="Budget vs Actual" \
#   --prop data="Budget:5000,Sales:2000,Marketing:-800,Ops:-600,Net:5600" \
#   --prop increaseColor=2E75B6 \
#   --prop decreaseColor=C00000 \
#   --prop totalColor=FFC000 \
#   --prop legend=bottom \
#   --prop chartFill=F0F4FA \
#   --prop dataLabels=true \
#   --prop labelFont="9:333333:true"
#
# Features: waterfall with legend=bottom, chartFill (solid hex — cx charts
#   don't support gradient fills, use plain RGB), labelFont "size:color:bold"
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Waterfall & Funnel" --type chart'
    f' --prop chartType=waterfall'
    f' --prop title="Budget vs Actual"'
    f' --prop data=Budget:5000,Sales:2000,Marketing:-800,Ops:-600,Net:5600'
    f' --prop increaseColor=2E75B6'
    f' --prop decreaseColor=C00000'
    f' --prop totalColor=FFC000'
    f' --prop legend=bottom'
    f' --prop chartFill=F0F4FA'
    f' --prop dataLabels=true'
    f' --prop labelFont=9:333333:true'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 3: Funnel — sales pipeline with title shadow
#
# officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
#   --prop chartType=funnel \
#   --prop title="Sales Pipeline" \
#   --prop series1="Pipeline:1200,850,600,300,120" \
#   --prop categories=Leads,Qualified,Proposal,Negotiation,Won \
#   --prop dataLabels=true \
#   --prop title.shadow="000000-4-45-2-40"
#
# Features: chartType=funnel, descending pipeline values, dataLabels,
#   title.shadow "COLOR-BLUR-ANGLE-DIST-OPACITY"
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Waterfall & Funnel" --type chart'
    f' --prop chartType=funnel'
    f' --prop title="Sales Pipeline"'
    f' --prop series1=Pipeline:1200,850,600,300,120'
    f' --prop categories=Leads,Qualified,Proposal,Negotiation,Won'
    f' --prop dataLabels=true'
    f' --prop title.shadow=000000-4-45-2-40'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 4: Funnel — marketing conversion + legend/axis fonts + axis titles
#
# officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
#   --prop chartType=funnel \
#   --prop title="Marketing Funnel" \
#   --prop series1="Users:10000,6500,3200,1800,900,450" \
#   --prop categories=Impressions,Clicks,Signups,Active,Paying,Retained \
#   --prop dataLabels=true \
#   --prop legendfont="9:8B949E:Helvetica Neue" \
#   --prop axisfont="10:58626E:Helvetica Neue"
#
# Features: funnel, legendfont "size:color:fontname", axisfont,
#   6-stage pipeline, dataLabels
#
# NOTE: `colors=` palette is intentionally omitted here. On cx:chart single-
#   series types (funnel/treemap/sunburst) the CLI only applies the first
#   palette color to the whole series, so all bars would render the same
#   color. Let Excel's theme pick the default accent color.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Waterfall & Funnel" --type chart'
    f' --prop chartType=funnel'
    f' --prop title="Marketing Funnel"'
    f' --prop series1=Users:10000,6500,3200,1800,900,450'
    f' --prop categories=Impressions,Clicks,Signups,Active,Paying,Retained'
    f' --prop dataLabels=true'
    f' --prop "legendfont=9:8B949E:Helvetica Neue"'
    f' --prop "axisfont=10:58626E:Helvetica Neue"'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# ==========================================================================
# Sheet 2: Treemap & Sunburst
# ==========================================================================
print("\n--- 2-Treemap & Sunburst ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Treemap & Sunburst"')

# --------------------------------------------------------------------------
# Chart 1: Treemap — parentLabelLayout=overlapping + dataLabels
#
# officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
#   --prop chartType=treemap \
#   --prop title="Revenue by Product" \
#   --prop series1="Revenue:450,380,310,280,210,180,150,120" \
#   --prop categories=Laptops,Phones,Tablets,TVs,Cameras,Audio,Gaming,Wearables \
#   --prop parentLabelLayout=overlapping \
#   --prop dataLabels=true
#
# Features: chartType=treemap, parentLabelLayout=overlapping, dataLabels.
#   NOTE: `colors=` is omitted — see Funnel Chart 4 note: cx single-series
#   charts only pick up the first palette color. Excel's theme will auto-
#   rainbow the tiles instead.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Treemap & Sunburst" --type chart'
    f' --prop chartType=treemap'
    f' --prop title="Revenue by Product"'
    f' --prop series1=Revenue:450,380,310,280,210,180,150,120'
    f' --prop categories=Laptops,Phones,Tablets,TVs,Cameras,Audio,Gaming,Wearables'
    f' --prop parentLabelLayout=overlapping'
    f' --prop dataLabels=true'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 2: Treemap — parentLabelLayout=banner + bold title
#
# officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
#   --prop chartType=treemap \
#   --prop title="Department Budget" \
#   --prop series1="Budget:900,750,600,500,420,350,280" \
#   --prop categories=Engineering,Sales,Marketing,Support,Finance,HR,Legal \
#   --prop parentLabelLayout=banner \
#   --prop title.bold=true \
#   --prop title.size=14 \
#   --prop title.color=2E5090
#
# Features: treemap parentLabelLayout=banner, title.bold/size/color
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Treemap & Sunburst" --type chart'
    f' --prop chartType=treemap'
    f' --prop title="Department Budget"'
    f' --prop series1=Budget:900,750,600,500,420,350,280'
    f' --prop categories=Engineering,Sales,Marketing,Support,Finance,HR,Legal'
    f' --prop parentLabelLayout=banner'
    f' --prop title.bold=true'
    f' --prop title.size=14'
    f' --prop title.color=2E5090'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 3: Treemap — parentLabelLayout=none (no parent label strip)
#
# officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
#   --prop chartType=treemap \
#   --prop title="Flat Treemap (no parent labels)" \
#   --prop series1="Units:250,200,180,160,140,120,100,80,60,40" \
#   --prop categories=A,B,C,D,E,F,G,H,I,J \
#   --prop parentLabelLayout=none \
#   --prop dataLabels=true
#
# Features: treemap parentLabelLayout=none (all labels inline, no header strip),
#   dataLabels on leaf tiles
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Treemap & Sunburst" --type chart'
    f' --prop chartType=treemap'
    f' --prop title="Flat Treemap (no parent labels)"'
    f' --prop series1=Units:250,200,180,160,140,120,100,80,60,40'
    f' --prop categories=A,B,C,D,E,F,G,H,I,J'
    f' --prop parentLabelLayout=none'
    f' --prop dataLabels=true'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 4: Sunburst — radial hierarchy + chartFill (solid) + plotFill
#
# officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
#   --prop chartType=sunburst \
#   --prop title="Market Share by Region" \
#   --prop series1="Share:35,25,20,15,30,25,20,10,15" \
#   --prop categories=North,South,East,West,Urban,Suburban,Rural,Online,Retail \
#   --prop chartFill=F8FAFC \
#   --prop plotFill=FFFFFF \
#   --prop dataLabels=true
#
# Features: chartType=sunburst, radial hierarchical layout, chartFill (solid hex),
#   plotFill (solid hex), dataLabels.
#   NOTE 1: cx:chart's chart/plot fill only accepts solid color — not gradient
#     (unlike regular cChart). Use a single hex like "F8FAFC" or "none".
#   NOTE 2: `colors=` palette is omitted for the same reason as the funnel/
#     treemap examples — cx single-series charts paint only the first palette
#     entry. Let Excel's theme drive per-segment coloring.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Treemap & Sunburst" --type chart'
    f' --prop chartType=sunburst'
    f' --prop title="Market Share by Region"'
    f' --prop series1=Share:35,25,20,15,30,25,20,10,15'
    f' --prop categories=North,South,East,West,Urban,Suburban,Rural,Online,Retail'
    f' --prop chartFill=F8FAFC'
    f' --prop plotFill=FFFFFF'
    f' --prop dataLabels=true'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# ==========================================================================
# Sheet 3: Histogram & Box Whisker
# ==========================================================================
print("\n--- 3-Histogram & BoxWhisker ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Histogram & BoxWhisker"')

# --------------------------------------------------------------------------
# Chart 1: Histogram — auto-binning (Excel picks bin count)
#
# officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
#   --prop chartType=histogram \
#   --prop title="Test Scores (auto bins)" \
#   --prop series1="Scores:45,52,58,61,63,65,67,68,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,97,99"
#
# Features: chartType=histogram, no binning knobs → Excel auto-selects bins
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Histogram & BoxWhisker" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Test Scores (auto bins)"'
    f' --prop series1=Scores:45,52,58,61,63,65,67,68,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,97,99'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 2: Histogram — explicit binCount=5 with title glow
#
# officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
#   --prop chartType=histogram \
#   --prop title="Sales (binCount=5)" \
#   --prop series1="Sales:120,135,148,155,162,170,175,183,191,200,210,220,235,250,265,280,295,310,340,380,420,480,550,620,700" \
#   --prop binCount=5 \
#   --prop title.glow="FFC000-6-50"
#
# Features: histogram binCount (explicit bin count), title.glow
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Histogram & BoxWhisker" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Sales (binCount=5)"'
    f' --prop series1=Sales:120,135,148,155,162,170,175,183,191,200,210,220,235,250,265,280,295,310,340,380,420,480,550,620,700'
    f' --prop binCount=5'
    f' --prop title.glow=FFC000-6-50'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 3: Histogram — explicit binSize=50 (fixed bin width) + label font
#
# officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
#   --prop chartType=histogram \
#   --prop title="Sales (binSize=50)" \
#   --prop series1="Sales:120,135,148,155,162,170,175,183,191,200,210,220,235,250,265,280,295,310,340,380,420,480,550,620,700" \
#   --prop binSize=50 \
#   --prop dataLabels=true \
#   --prop labelFont="9:FFFFFF:true"
#
# Features: histogram binSize (explicit bin width — mutually exclusive with
#   binCount), dataLabels, labelFont
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Histogram & BoxWhisker" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Sales (binSize=50)"'
    f' --prop series1=Sales:120,135,148,155,162,170,175,183,191,200,210,220,235,250,265,280,295,310,340,380,420,480,550,620,700'
    f' --prop binSize=50'
    f' --prop dataLabels=true'
    f' --prop labelFont=9:FFFFFF:true'
    f' --prop x=28 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 4: Histogram — overflow/underflow bins + intervalClosed=l
#
# officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
#   --prop chartType=histogram \
#   --prop title="Response Time (outlier bins)" \
#   --prop series1="ms:40,55,68,75,82,88,95,102,110,118,125,135,150,175,220,280,350" \
#   --prop underflowBin=60 \
#   --prop overflowBin=200 \
#   --prop intervalClosed=l \
#   --prop dataLabels=true \
#   --prop legend=none
#
# Features: histogram underflowBin (cutoff for <N), overflowBin (cutoff for >N),
#   intervalClosed=l (bins are [a,b) — left-closed; default "r" is (a,b]),
#   legend=none
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Histogram & BoxWhisker" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Response Time (outlier bins)"'
    f' --prop series1=ms:40,55,68,75,82,88,95,102,110,118,125,135,150,175,220,280,350'
    f' --prop underflowBin=60'
    f' --prop overflowBin=200'
    f' --prop intervalClosed=l'
    f' --prop dataLabels=true'
    f' --prop legend=none'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 5: Box & Whisker — two teams, quartileMethod=exclusive
#
# officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
#   --prop chartType=boxWhisker \
#   --prop title="Response Time by Team (ms)" \
#   --prop series1="TeamA:42,55,61,68,72,75,78,81,85,88,92,97,105,120" \
#   --prop series2="TeamB:30,38,45,52,58,62,65,68,71,74,78,85,92,110" \
#   --prop quartileMethod=exclusive \
#   --prop legend=bottom
#
# Features: chartType=boxWhisker, two-series comparison,
#   quartileMethod=exclusive, legend=bottom, outlier detection (built-in)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Histogram & BoxWhisker" --type chart'
    f' --prop chartType=boxWhisker'
    f' --prop title="Response Time by Team (ms)"'
    f' --prop "series1=TeamA:42,55,61,68,72,75,78,81,85,88,92,97,105,120"'
    f' --prop "series2=TeamB:30,38,45,52,58,62,65,68,71,74,78,85,92,110"'
    f' --prop quartileMethod=exclusive'
    f' --prop legend=bottom'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 6: Box & Whisker — three departments, quartileMethod=inclusive + title glow
#
# officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
#   --prop chartType=boxWhisker \
#   --prop title="Salary Distribution ($k)" \
#   --prop series1="Engineering:85,92,95,98,102,105,108,112,118,125,135,150,180" \
#   --prop series2="Marketing:60,65,68,72,75,78,80,83,88,92,98,110" \
#   --prop series3="Sales:55,62,68,75,82,90,98,105,115,125,140,160,190" \
#   --prop quartileMethod=inclusive \
#   --prop title.glow="00D2FF-6-60" \
#   --prop legend=bottom
#
# Features: boxWhisker three-series, quartileMethod=inclusive (different
#   quartile formula from exclusive), title.glow, mean markers (default on)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Histogram & BoxWhisker" --type chart'
    f' --prop chartType=boxWhisker'
    f' --prop title="Salary Distribution (\\$k)"'
    f' --prop "series1=Engineering:85,92,95,98,102,105,108,112,118,125,135,150,180"'
    f' --prop "series2=Marketing:60,65,68,72,75,78,80,83,88,92,98,110"'
    f' --prop "series3=Sales:55,62,68,75,82,90,98,105,115,125,140,160,190"'
    f' --prop quartileMethod=inclusive'
    f' --prop title.glow=00D2FF-6-60'
    f' --prop legend=bottom'
    f' --prop x=28 --prop y=19 --prop width=13 --prop height=18')

# ==========================================================================
# Sheet 4: Pareto
# ==========================================================================
print("\n--- 4-Pareto ---")
cli(f'add "{FILE}" / --type sheet --prop name="4-Pareto"')

# --------------------------------------------------------------------------
# Chart 1: Pareto — defect analysis, raw counts auto-sorted + cumul% overlay
#
# officecli add charts-extended.xlsx "/4-Pareto" --type chart \
#   --prop chartType=pareto \
#   --prop title="Defect Pareto" \
#   --prop series1="Count:45,30,10,8,5,2" \
#   --prop categories=Scratches,Dents,Cracks,Chips,Stains,Other \
#   --prop dataLabels=true
#
# Features: chartType=pareto (2-series under the hood — clusteredColumn bars
#   + paretoLine cumulative %), automatic descending sort, cumulative %
#   computed server-side, dataLabels on both series.
#   Input is a SINGLE user series; officecli pre-sorts by value desc and
#   emits the two cx:series MSO expects (layoutId=clusteredColumn +
#   layoutId=paretoLine with cx:binning intervalClosed="r").
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Pareto" --type chart'
    f' --prop chartType=pareto'
    f' --prop title="Defect Pareto"'
    f' --prop series1=Count:45,30,10,8,5,2'
    f' --prop categories=Scratches,Dents,Cracks,Chips,Stains,Other'
    f' --prop dataLabels=true'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 2: Pareto — root cause analysis, 10 categories, out-of-order input
#
# officecli add charts-extended.xlsx "/4-Pareto" --type chart \
#   --prop chartType=pareto \
#   --prop title="Root Cause Pareto" \
#   --prop series1="Tickets:12,87,5,45,3,120,22,67,8,31" \
#   --prop categories=Network,Auth,DB,Cache,UI,Config,Deploy,Monitor,Queue,Storage \
#   --prop title.glow="FFC000-6-50" \
#   --prop legend=bottom
#
# Features: pareto with unsorted input values (12, 87, 5, ...) — officecli
#   re-sorts by value desc (120, 87, 67, ...) and re-aligns categories so
#   the biggest contributor renders first. title.glow + legend=bottom
#   demonstrate generic cx styling on pareto.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Pareto" --type chart'
    f' --prop chartType=pareto'
    f' --prop title="Root Cause Pareto"'
    f' --prop series1=Tickets:12,87,5,45,3,120,22,67,8,31'
    f' --prop categories=Network,Auth,DB,Cache,UI,Config,Deploy,Monitor,Queue,Storage'
    f' --prop title.glow=FFC000-6-50'
    f' --prop legend=bottom'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# Remove blank default Sheet1 (all data is inline)
cli(f'remove "{FILE}" /Sheet1')

print(f"\nDone! Generated: {FILE}")
print("  4 sheets, 16 charts total (full cx:chart feature coverage)")
print("  Sheet 1: Waterfall (2) + Funnel (2)")
print("  Sheet 2: Treemap (3: overlapping/banner/none) + Sunburst (1)")
print("  Sheet 3: Histogram (4: auto/binCount/binSize/overflow+underflow+intervalClosed=l) + BoxWhisker (2: exclusive/inclusive)")
print("  Sheet 4: Pareto (2: sorted input / out-of-order input)")
