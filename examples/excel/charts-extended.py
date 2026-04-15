#!/usr/bin/env python3
"""
Extended Chart Types Showcase — waterfall, funnel, treemap, sunburst, histogram, boxWhisker.

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
# Sheet: 1-Waterfall & Funnel
# ==========================================================================
print("\n--- 1-Waterfall & Funnel ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Waterfall & Funnel"')

# --------------------------------------------------------------------------
# Chart 1: Waterfall with increase/decrease/total colors and data labels
#
# officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
#   --prop chartType=waterfall \
#   --prop title="Cash Flow Bridge" \
#   --prop data="Start:1000,Revenue:500,Costs:-300,Tax:-100,Net:1100" \
#   --prop increaseColor=70AD47 \
#   --prop decreaseColor=FF0000 \
#   --prop totalColor=4472C4 \
#   --prop x=0 --prop y=0 --prop width=13 --prop height=18 \
#   --prop dataLabels=true
#
# Features: chartType=waterfall, data= name:value pairs, increaseColor,
#   decreaseColor, totalColor, dataLabels
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Waterfall & Funnel" --type chart'
    f' --prop chartType=waterfall'
    f' --prop title="Cash Flow Bridge"'
    f' --prop data=Start:1000,Revenue:500,Costs:-300,Tax:-100,Net:1100'
    f' --prop increaseColor=70AD47'
    f' --prop decreaseColor=FF0000'
    f' --prop totalColor=4472C4'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18'
    f' --prop dataLabels=true')

# --------------------------------------------------------------------------
# Chart 2: Waterfall with custom color theme and explicit categories
#
# officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
#   --prop chartType=waterfall \
#   --prop title="Budget vs Actual" \
#   --prop data="Budget:5000,Sales:2000,Marketing:-800,Ops:-600,Net:5600" \
#   --prop increaseColor=2E75B6 \
#   --prop decreaseColor=C00000 \
#   --prop totalColor=FFC000 \
#   --prop x=14 --prop y=0 --prop width=13 --prop height=18 \
#   --prop legend=bottom
#
# Features: waterfall legend, alternative color palette (blue/red/amber)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Waterfall & Funnel" --type chart'
    f' --prop chartType=waterfall'
    f' --prop title="Budget vs Actual"'
    f' --prop data=Budget:5000,Sales:2000,Marketing:-800,Ops:-600,Net:5600'
    f' --prop increaseColor=2E75B6'
    f' --prop decreaseColor=C00000'
    f' --prop totalColor=FFC000'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Funnel chart — sales pipeline with descending values
#
# officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
#   --prop chartType=funnel \
#   --prop title="Sales Pipeline" \
#   --prop series1="Pipeline:1200,850,600,300,120" \
#   --prop categories=Leads,Qualified,Proposal,Negotiation,Won \
#   --prop x=0 --prop y=19 --prop width=13 --prop height=18 \
#   --prop dataLabels=true
#
# Features: chartType=funnel, descending values (pipeline stages), dataLabels
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Waterfall & Funnel" --type chart'
    f' --prop chartType=funnel'
    f' --prop title="Sales Pipeline"'
    f' --prop series1=Pipeline:1200,850,600,300,120'
    f' --prop categories=Leads,Qualified,Proposal,Negotiation,Won'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18'
    f' --prop dataLabels=true')

# --------------------------------------------------------------------------
# Chart 4: Funnel chart — conversion funnel with data labels
#
# officecli add charts-extended.xlsx "/1-Waterfall & Funnel" --type chart \
#   --prop chartType=funnel \
#   --prop title="Marketing Funnel" \
#   --prop series1="Users:10000,6500,3200,1800,900,450" \
#   --prop categories=Impressions,Clicks,Signups,Active,Paying,Retained \
#   --prop x=14 --prop y=19 --prop width=13 --prop height=18 \
#   --prop dataLabels=true
#
# Features: funnel with 6 stages, large-range values, dataLabels
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Waterfall & Funnel" --type chart'
    f' --prop chartType=funnel'
    f' --prop title="Marketing Funnel"'
    f' --prop series1=Users:10000,6500,3200,1800,900,450'
    f' --prop categories=Impressions,Clicks,Signups,Active,Paying,Retained'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18'
    f' --prop dataLabels=true')

# ==========================================================================
# Sheet: 2-Treemap & Sunburst
# ==========================================================================
print("\n--- 2-Treemap & Sunburst ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Treemap & Sunburst"')

# --------------------------------------------------------------------------
# Chart 1: Treemap — product revenue by category
#
# officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
#   --prop chartType=treemap \
#   --prop title="Revenue by Product" \
#   --prop series1="Revenue:450,380,310,280,210,180,150,120" \
#   --prop categories=Laptops,Phones,Tablets,TVs,Cameras,Audio,Gaming,Wearables \
#   --prop x=0 --prop y=0 --prop width=13 --prop height=18 \
#   --prop parentLabelLayout=overlapping
#
# Features: chartType=treemap, parentLabelLayout=overlapping, area-sized blocks
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Treemap & Sunburst" --type chart'
    f' --prop chartType=treemap'
    f' --prop title="Revenue by Product"'
    f' --prop series1=Revenue:450,380,310,280,210,180,150,120'
    f' --prop categories=Laptops,Phones,Tablets,TVs,Cameras,Audio,Gaming,Wearables'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18'
    f' --prop parentLabelLayout=overlapping')

# --------------------------------------------------------------------------
# Chart 2: Treemap — budget allocation with banner labels
#
# officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
#   --prop chartType=treemap \
#   --prop title="Department Budget" \
#   --prop series1="Budget:900,750,600,500,420,350,280" \
#   --prop categories=Engineering,Sales,Marketing,Support,Finance,HR,Legal \
#   --prop x=14 --prop y=0 --prop width=13 --prop height=18 \
#   --prop parentLabelLayout=banner
#
# Features: treemap parentLabelLayout=banner (header strip style)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Treemap & Sunburst" --type chart'
    f' --prop chartType=treemap'
    f' --prop title="Department Budget"'
    f' --prop series1=Budget:900,750,600,500,420,350,280'
    f' --prop categories=Engineering,Sales,Marketing,Support,Finance,HR,Legal'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18'
    f' --prop parentLabelLayout=banner')

# --------------------------------------------------------------------------
# Chart 3: Sunburst — hierarchical market breakdown
#
# officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
#   --prop chartType=sunburst \
#   --prop title="Market Share by Region" \
#   --prop series1="Share:35,25,20,15,30,25,20,10,15" \
#   --prop categories=North,South,East,West,Urban,Suburban,Rural,Online,Retail \
#   --prop x=0 --prop y=19 --prop width=13 --prop height=18
#
# Features: chartType=sunburst, radial hierarchical layout
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Treemap & Sunburst" --type chart'
    f' --prop chartType=sunburst'
    f' --prop title="Market Share by Region"'
    f' --prop series1=Share:35,25,20,15,30,25,20,10,15'
    f' --prop categories=North,South,East,West,Urban,Suburban,Rural,Online,Retail'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 4: Sunburst — product category breakdown
#
# officecli add charts-extended.xlsx "/2-Treemap & Sunburst" --type chart \
#   --prop chartType=sunburst \
#   --prop title="Product Categories" \
#   --prop series1="Units:500,400,300,250,200,180,160,140,120,100" \
#   --prop categories=A,B,C,D,E,F,G,H,I,J \
#   --prop x=14 --prop y=19 --prop width=13 --prop height=18
#
# Features: sunburst with 10 segments, unit-count values
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Treemap & Sunburst" --type chart'
    f' --prop chartType=sunburst'
    f' --prop title="Product Categories"'
    f' --prop series1=Units:500,400,300,250,200,180,160,140,120,100'
    f' --prop categories=A,B,C,D,E,F,G,H,I,J'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# ==========================================================================
# Sheet: 3-Histogram & Box Whisker
# ==========================================================================
print("\n--- 3-Histogram & Box Whisker ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Histogram & BoxWhisker"')

# --------------------------------------------------------------------------
# Chart 1: Histogram — test score distribution with auto-binning
#
# officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
#   --prop chartType=histogram \
#   --prop title="Test Score Distribution" \
#   --prop series1="Scores:45,52,58,61,63,65,67,68,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,97,99" \
#   --prop x=0 --prop y=0 --prop width=13 --prop height=18
#
# Features: chartType=histogram, raw value distribution, auto-binning
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Histogram & BoxWhisker" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Test Score Distribution"'
    f' --prop series1=Scores:45,52,58,61,63,65,67,68,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,97,99'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 2: Histogram — sales amounts with explicit bin count
#
# officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
#   --prop chartType=histogram \
#   --prop title="Sales Distribution (5 bins)" \
#   --prop series1="Sales:120,135,148,155,162,170,175,183,191,200,210,220,235,250,265,280,295,310,340,380,420,480,550,620,700" \
#   --prop binCount=5 \
#   --prop x=14 --prop y=0 --prop width=13 --prop height=18
#
# Features: histogram with binCount=5 (explicit bin count)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Histogram & BoxWhisker" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Sales Distribution (5 bins)"'
    f' --prop series1=Sales:120,135,148,155,162,170,175,183,191,200,210,220,235,250,265,280,295,310,340,380,420,480,550,620,700'
    f' --prop binCount=5'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 3: Box & Whisker — response time comparison across teams
#
# officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
#   --prop chartType=boxWhisker \
#   --prop title="Response Time by Team (ms)" \
#   --prop series1="TeamA:42,55,61,68,72,75,78,81,85,88,92,97,105,120" \
#   --prop series2="TeamB:30,38,45,52,58,62,65,68,71,74,78,85,92,110" \
#   --prop x=0 --prop y=19 --prop width=13 --prop height=18 \
#   --prop quartileMethod=exclusive
#
# Features: chartType=boxWhisker, two series (grouped), quartileMethod=exclusive,
#   outlier detection
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Histogram & BoxWhisker" --type chart'
    f' --prop chartType=boxWhisker'
    f' --prop title="Response Time by Team (ms)"'
    f' --prop "series1=TeamA:42,55,61,68,72,75,78,81,85,88,92,97,105,120"'
    f' --prop "series2=TeamB:30,38,45,52,58,62,65,68,71,74,78,85,92,110"'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18'
    f' --prop quartileMethod=exclusive')

# --------------------------------------------------------------------------
# Chart 4: Box & Whisker — salary distribution across departments
#
# officecli add charts-extended.xlsx "/3-Histogram & BoxWhisker" --type chart \
#   --prop chartType=boxWhisker \
#   --prop title="Salary Distribution ($k)" \
#   --prop series1="Engineering:85,92,95,98,102,105,108,112,118,125,135,150,180" \
#   --prop series2="Marketing:60,65,68,72,75,78,80,83,88,92,98,110" \
#   --prop series3="Sales:55,62,68,75,82,90,98,105,115,125,140,160,190" \
#   --prop x=14 --prop y=19 --prop width=13 --prop height=18 \
#   --prop quartileMethod=inclusive
#
# Features: boxWhisker three-series, quartileMethod=inclusive, mean markers
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Histogram & BoxWhisker" --type chart'
    f' --prop chartType=boxWhisker'
    f' --prop title="Salary Distribution (\\$k)"'
    f' --prop "series1=Engineering:85,92,95,98,102,105,108,112,118,125,135,150,180"'
    f' --prop "series2=Marketing:60,65,68,72,75,78,80,83,88,92,98,110"'
    f' --prop "series3=Sales:55,62,68,75,82,90,98,105,115,125,140,160,190"'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18'
    f' --prop quartileMethod=inclusive')

# Remove blank default Sheet1 (all data is inline)
cli(f'remove "{FILE}" /Sheet1')

print(f"\nDone! Generated: {FILE}")
print("  3 sheets (12 charts total)")
print("  Sheet 1: Waterfall (2) + Funnel (2)")
print("  Sheet 2: Treemap (2) + Sunburst (2)")
print("  Sheet 3: Histogram (2) + BoxWhisker (2)")
