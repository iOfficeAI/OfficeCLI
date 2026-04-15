#!/usr/bin/env python3
"""
Bar (Horizontal) Charts Showcase — bar, barStacked, barPercentStacked, and bar3d with all variations.

Generates: charts-bar.xlsx

Every horizontal bar chart feature officecli supports is demonstrated at least once:
gap width, overlap, data labels, axis scaling, gridlines, legend positioning,
reference lines, secondary axis, error bars, gradients, transparency, shadows,
manual layout, data table, 3D rotation, and conditional coloring.

6 sheets, 24 charts total.

  1-Bar Fundamentals      4 charts — data input variants, colors, stacked, data shorthand
  2-Bar Variants          4 charts — barStacked, barPercentStacked, bar3d, cylinder
  3-Bar Styling           4 charts — title styling, shadow/outline, gradients, plot/chart fill
  4-Axis & Labels         4 charts — axis scale, log/reverse/dispUnits, label styling, per-point
  5-Legend & Layout        4 charts — legend positions, overlay, manual layout, secondary axis
  6-Advanced              4 charts — reference line, colorRule, glow/shadow, errBars/dataTable

Usage:
  python3 charts-bar.py
"""

import subprocess, sys, os, json, atexit

FILE = "charts-bar.xlsx"

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
for j, h in enumerate(["Department", "Q1", "Q2", "Q3", "Q4"]):
    data_cmds.append({"command": "set", "path": f"/Sheet1/{'ABCDE'[j]}1", "props": {"text": h, "bold": "true"}})

depts = ["Engineering", "Marketing", "Sales", "Support", "Finance", "HR", "Legal", "Operations"]
q1 =    [185, 120, 210, 95, 78, 62, 55, 140]
q2 =    [195, 135, 225, 105, 82, 68, 58, 152]
q3 =    [210, 142, 240, 112, 88, 72, 62, 165]
q4 =    [228, 158, 260, 118, 92, 78, 68, 178]

for i in range(8):
    r = i + 2
    for j, val in enumerate([depts[i], q1[i], q2[i], q3[i], q4[i]]):
        data_cmds.append({"command": "set", "path": f"/Sheet1/{'ABCDE'[j]}{r}", "props": {"text": str(val)}})

cli(f'batch "{FILE}" --force --commands \'{json.dumps(data_cmds)}\'')

# ==========================================================================
# Sheet: 1-Bar Fundamentals
# ==========================================================================
print("\n--- 1-Bar Fundamentals ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Bar Fundamentals"')

# --------------------------------------------------------------------------
# Chart 1: Basic bar chart with dataRange, axis titles, and gridlines
#
# officecli add charts-bar.xlsx "/1-Bar Fundamentals" --type chart \
#   --prop chartType=bar \
#   --prop title="Department Performance — Q1" \
#   --prop dataRange=Sheet1!A1:B9 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop catTitle=Department --prop axisTitle=Score \
#   --prop axisfont=9:333333:Arial \
#   --prop gridlines=D9D9D9:0.5:dot
#
# Features: chartType=bar, dataRange, catTitle, axisTitle, axisfont, gridlines
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Bar Fundamentals" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Department Performance — Q1"'
    f' --prop dataRange=Sheet1!A1:B9'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop catTitle=Department --prop axisTitle=Score'
    f' --prop axisfont=9:333333:Arial'
    f' --prop gridlines=D9D9D9:0.5:dot')

# --------------------------------------------------------------------------
# Chart 2: Inline series with custom colors, gap width, and data labels
#
# officecli add charts-bar.xlsx "/1-Bar Fundamentals" --type chart \
#   --prop chartType=bar \
#   --prop title="Survey Results" \
#   --prop series1="Satisfaction:85,72,91,68,78" \
#   --prop categories=Product,Service,Delivery,Price,Overall \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000,5B9BD5 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop gapwidth=80 \
#   --prop dataLabels=outsideEnd
#
# Features: inline series, colors per category, gapwidth, dataLabels=outsideEnd
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Bar Fundamentals" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Survey Results"'
    f' --prop series1=Satisfaction:85,72,91,68,78'
    f' --prop categories=Product,Service,Delivery,Price,Overall'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000,5B9BD5'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop gapwidth=80'
    f' --prop dataLabels=outsideEnd')

# --------------------------------------------------------------------------
# Chart 3: Stacked bar with overlap and series outline
#
# officecli add charts-bar.xlsx "/1-Bar Fundamentals" --type chart \
#   --prop chartType=barStacked \
#   --prop title="Quarterly Headcount by Dept" \
#   --prop series1="Q1:30,18,25,12" \
#   --prop series2="Q2:35,20,28,14" \
#   --prop series3="Q3:38,22,30,16" \
#   --prop categories=Engineering,Marketing,Sales,Support \
#   --prop colors=2E75B6,70AD47,FFC000 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop overlap=0 \
#   --prop series.outline=FFFFFF-0.5
#
# Features: barStacked, overlap=0, series.outline (white separator)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Bar Fundamentals" --type chart'
    f' --prop chartType=barStacked'
    f' --prop title="Quarterly Headcount by Dept"'
    f' --prop series1=Q1:30,18,25,12'
    f' --prop series2=Q2:35,20,28,14'
    f' --prop series3=Q3:38,22,30,16'
    f' --prop categories=Engineering,Marketing,Sales,Support'
    f' --prop colors=2E75B6,70AD47,FFC000'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop overlap=0'
    f' --prop series.outline=FFFFFF-0.5')

# --------------------------------------------------------------------------
# Chart 4: data= shorthand with legend=bottom
#
# officecli add charts-bar.xlsx "/1-Bar Fundamentals" --type chart \
#   --prop chartType=bar \
#   --prop title="Training Hours by Team" \
#   --prop 'data=Technical:45,38,52;Soft Skills:20,28,18;Compliance:12,15,10' \
#   --prop categories=Engineering,Sales,Support \
#   --prop colors=4472C4,ED7D31,70AD47 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop legend=bottom
#
# Features: data= shorthand (inline multi-series), legend=bottom
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Bar Fundamentals" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Training Hours by Team"'
    f' --prop "data=Technical:45,38,52;Soft Skills:20,28,18;Compliance:12,15,10"'
    f' --prop categories=Engineering,Sales,Support'
    f' --prop colors=4472C4,ED7D31,70AD47'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 2-Bar Variants
# ==========================================================================
print("\n--- 2-Bar Variants ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Bar Variants"')

# --------------------------------------------------------------------------
# Chart 1: barStacked with tight gap width
#
# officecli add charts-bar.xlsx "/2-Bar Variants" --type chart \
#   --prop chartType=barStacked \
#   --prop title="Budget Allocation" \
#   --prop series1="Salaries:120,80,95,60" \
#   --prop series2="Operations:45,35,40,25" \
#   --prop series3="Marketing:30,50,20,15" \
#   --prop categories=Engineering,Sales,Support,HR \
#   --prop colors=1F4E79,2E75B6,9DC3E6 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop gapwidth=50 \
#   --prop legend=bottom
#
# Features: barStacked, gapwidth=50 (tight bars)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bar Variants" --type chart'
    f' --prop chartType=barStacked'
    f' --prop title="Budget Allocation"'
    f' --prop series1=Salaries:120,80,95,60'
    f' --prop series2=Operations:45,35,40,25'
    f' --prop series3=Marketing:30,50,20,15'
    f' --prop categories=Engineering,Sales,Support,HR'
    f' --prop colors=1F4E79,2E75B6,9DC3E6'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop gapwidth=50'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 2: barPercentStacked with axis number format and reference line
#
# officecli add charts-bar.xlsx "/2-Bar Variants" --type chart \
#   --prop chartType=barPercentStacked \
#   --prop title="Task Completion Ratio" \
#   --prop series1="Done:75,60,90,45,80" \
#   --prop series2="In Progress:15,25,5,30,12" \
#   --prop series3="Blocked:10,15,5,25,8" \
#   --prop categories=Backend,Frontend,QA,Design,DevOps \
#   --prop colors=70AD47,FFC000,C00000 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop axisNumFmt=0% \
#   --prop referenceLine=0.5:FF0000:Target:dash \
#   --prop legend=bottom
#
# Features: barPercentStacked, axisNumFmt=0%, referenceLine with label and dash
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bar Variants" --type chart'
    f' --prop chartType=barPercentStacked'
    f' --prop title="Task Completion Ratio"'
    f' --prop series1=Done:75,60,90,45,80'
    f' --prop series2="In Progress:15,25,5,30,12"'
    f' --prop series3=Blocked:10,15,5,25,8'
    f' --prop categories=Backend,Frontend,QA,Design,DevOps'
    f' --prop colors=70AD47,FFC000,C00000'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop axisNumFmt=0%'
    f' --prop referenceLine=0.5:FF0000:Target:dash'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: bar3d with perspective and style
#
# officecli add charts-bar.xlsx "/2-Bar Variants" --type chart \
#   --prop chartType=bar3d \
#   --prop title="3D Revenue by Region" \
#   --prop series1="Revenue:340,280,310,195" \
#   --prop categories=North,South,East,West \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop view3d=10,30,20 \
#   --prop style=3 \
#   --prop legend=right
#
# Features: bar3d, view3d (rotX,rotY,perspective), style=3
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bar Variants" --type chart'
    f' --prop chartType=bar3d'
    f' --prop title="3D Revenue by Region"'
    f' --prop series1=Revenue:340,280,310,195'
    f' --prop categories=North,South,East,West'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop view3d=10,30,20'
    f' --prop style=3'
    f' --prop legend=right')

# --------------------------------------------------------------------------
# Chart 4: bar3d with cylinder shape
#
# officecli add charts-bar.xlsx "/2-Bar Variants" --type chart \
#   --prop chartType=bar3d \
#   --prop title="Cylinder — Project Milestones" \
#   --prop series1="Completed:8,12,6,10,15" \
#   --prop series2="Remaining:4,3,6,5,2" \
#   --prop categories=Alpha,Beta,Gamma,Delta,Epsilon \
#   --prop colors=2E75B6,BDD7EE \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop shape=cylinder \
#   --prop gapwidth=60 \
#   --prop legend=bottom
#
# Features: bar3d shape=cylinder, multi-series 3D bars
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Bar Variants" --type chart'
    f' --prop chartType=bar3d'
    f' --prop title="Cylinder — Project Milestones"'
    f' --prop series1=Completed:8,12,6,10,15'
    f' --prop series2=Remaining:4,3,6,5,2'
    f' --prop categories=Alpha,Beta,Gamma,Delta,Epsilon'
    f' --prop colors=2E75B6,BDD7EE'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop shape=cylinder'
    f' --prop gapwidth=60'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 3-Bar Styling
# ==========================================================================
print("\n--- 3-Bar Styling ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Bar Styling"')

# --------------------------------------------------------------------------
# Chart 1: Title styling (font, size, color, bold)
#
# officecli add charts-bar.xlsx "/3-Bar Styling" --type chart \
#   --prop chartType=bar \
#   --prop title="Styled Title Demo" \
#   --prop series1="Score:88,76,92,65,84" \
#   --prop categories=Dept A,Dept B,Dept C,Dept D,Dept E \
#   --prop colors=4472C4 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop title.font=Georgia --prop title.size=16 \
#   --prop title.color=1F4E79 --prop title.bold=true \
#   --prop gapwidth=100
#
# Features: title.font, title.size, title.color, title.bold
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Bar Styling" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Styled Title Demo"'
    f' --prop series1=Score:88,76,92,65,84'
    f' --prop categories=Dept A,Dept B,Dept C,Dept D,Dept E'
    f' --prop colors=4472C4'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop title.font=Georgia --prop title.size=16'
    f' --prop title.color=1F4E79 --prop title.bold=true'
    f' --prop gapwidth=100')

# --------------------------------------------------------------------------
# Chart 2: Series shadow and outline effects
#
# officecli add charts-bar.xlsx "/3-Bar Styling" --type chart \
#   --prop chartType=bar \
#   --prop title="Shadow & Outline" \
#   --prop series1="2024:165,142,180,128" \
#   --prop series2="2025:185,158,195,140" \
#   --prop categories=Engineering,Marketing,Sales,Support \
#   --prop colors=2E75B6,ED7D31 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop series.shadow=000000-4-315-2-30 \
#   --prop series.outline=1F4E79-1 \
#   --prop legend=bottom
#
# Features: series.shadow (color-blur-angle-dist-opacity), series.outline
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Bar Styling" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Shadow & Outline"'
    f' --prop series1=2024:165,142,180,128'
    f' --prop series2=2025:185,158,195,140'
    f' --prop categories=Engineering,Marketing,Sales,Support'
    f' --prop colors=2E75B6,ED7D31'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop series.shadow=000000-4-315-2-30'
    f' --prop series.outline=1F4E79-1'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 3: Per-series gradients
#
# officecli add charts-bar.xlsx "/3-Bar Styling" --type chart \
#   --prop chartType=bar \
#   --prop title="Gradient Bars" \
#   --prop series1="Revenue:320,275,410,190,245" \
#   --prop categories=North,South,East,West,Central \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop 'gradients=1F4E79-5B9BD5:0;C55A11-F4B183:0;548235-A9D18E:0;7F6000-FFD966:0;843C0B-DDA15E:0' \
#   --prop dataLabels=outsideEnd \
#   --prop labelFont=9:333333:true
#
# Features: gradients (per-bar gradient fills, angle=0 for horizontal),
#   labelFont (size:color:bold)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Bar Styling" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Gradient Bars"'
    f' --prop series1=Revenue:320,275,410,190,245'
    f' --prop categories=North,South,East,West,Central'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop "gradients=1F4E79-5B9BD5:0;C55A11-F4B183:0;548235-A9D18E:0;7F6000-FFD966:0;843C0B-DDA15E:0"'
    f' --prop dataLabels=outsideEnd'
    f' --prop labelFont=9:333333:true')

# --------------------------------------------------------------------------
# Chart 4: Plot fill gradient, chart fill, transparency, rounded corners
#
# officecli add charts-bar.xlsx "/3-Bar Styling" --type chart \
#   --prop chartType=bar \
#   --prop title="Styled Background" \
#   --prop dataRange=Sheet1!A1:C9 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop colors=5B9BD5,ED7D31 \
#   --prop plotFill=F0F4F8-D6E4F0:90 \
#   --prop chartFill=FFFFFF \
#   --prop transparency=20 \
#   --prop roundedCorners=true \
#   --prop legend=right
#
# Features: plotFill gradient, chartFill, transparency, roundedCorners
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Bar Styling" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Styled Background"'
    f' --prop dataRange=Sheet1!A1:C9'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop colors=5B9BD5,ED7D31'
    f' --prop plotFill=F0F4F8-D6E4F0:90'
    f' --prop chartFill=FFFFFF'
    f' --prop transparency=20'
    f' --prop roundedCorners=true'
    f' --prop legend=right')

# ==========================================================================
# Sheet: 4-Axis & Labels
# ==========================================================================
print("\n--- 4-Axis & Labels ---")
cli(f'add "{FILE}" / --type sheet --prop name="4-Axis & Labels"')

# --------------------------------------------------------------------------
# Chart 1: Custom axis min/max, majorUnit, and gridlines styling
#
# officecli add charts-bar.xlsx "/4-Axis & Labels" --type chart \
#   --prop chartType=bar \
#   --prop title="Axis Scale (50–250)" \
#   --prop dataRange=Sheet1!A1:B9 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop axisMin=50 --prop axisMax=250 --prop majorUnit=50 \
#   --prop gridlines=D0D0D0:0.5:solid \
#   --prop minorGridlines=EEEEEE:0.3:dot \
#   --prop axisLine=C00000:1.5:solid \
#   --prop catAxisLine=2E75B6:1.5:solid
#
# Features: axisMin, axisMax, majorUnit, gridlines styling,
#   minorGridlines, axisLine, catAxisLine
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Axis & Labels" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Axis Scale (50–250)"'
    f' --prop dataRange=Sheet1!A1:B9'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop axisMin=50 --prop axisMax=250 --prop majorUnit=50'
    f' --prop gridlines=D0D0D0:0.5:solid'
    f' --prop minorGridlines=EEEEEE:0.3:dot'
    f' --prop axisLine=C00000:1.5:solid'
    f' --prop catAxisLine=2E75B6:1.5:solid')

# --------------------------------------------------------------------------
# Chart 2: Log scale, axis reverse, and display units
#
# officecli add charts-bar.xlsx "/4-Axis & Labels" --type chart \
#   --prop chartType=bar \
#   --prop title="Log Scale & Reverse" \
#   --prop series1="Users:10,100,1000,5000,25000,100000" \
#   --prop categories=Tier 1,Tier 2,Tier 3,Tier 4,Tier 5,Tier 6 \
#   --prop colors=2E75B6 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop logBase=10 \
#   --prop axisReverse=true \
#   --prop dispUnits=thousands \
#   --prop gridlines=E0E0E0:0.5:dash
#
# Features: logBase=10, axisReverse=true, dispUnits=thousands
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Axis & Labels" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Log Scale & Reverse"'
    f' --prop "series1=Users:10,100,1000,5000,25000,100000"'
    f' --prop "categories=Tier 1,Tier 2,Tier 3,Tier 4,Tier 5,Tier 6"'
    f' --prop colors=2E75B6'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop logBase=10'
    f' --prop axisReverse=true'
    f' --prop dispUnits=thousands'
    f' --prop gridlines=E0E0E0:0.5:dash')

# --------------------------------------------------------------------------
# Chart 3: Data labels with labelFont, numFmt, separator
#
# officecli add charts-bar.xlsx "/4-Axis & Labels" --type chart \
#   --prop chartType=bar \
#   --prop title="Labeled Metrics" \
#   --prop series1="FY2025:148,92,215,178,125" \
#   --prop categories=Revenue,Costs,Gross,EBITDA,Net Income \
#   --prop colors=4472C4 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelPos=outsideEnd \
#   --prop labelFont=10:1F4E79:true \
#   --prop dataLabels.numFmt=#,##0 \
#   --prop "dataLabels.separator=: "
#
# Features: dataLabels, labelFont, dataLabels.numFmt, dataLabels.separator
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Axis & Labels" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Labeled Metrics"'
    f' --prop series1=FY2025:148,92,215,178,125'
    f' --prop categories=Revenue,Costs,Gross,EBITDA,Net Income'
    f' --prop colors=4472C4'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop dataLabels=outsideEnd'
    f' --prop labelFont=10:1F4E79:true'
    f' --prop dataLabels.numFmt=#,##0'
    f' --prop "dataLabels.separator=: "')

# --------------------------------------------------------------------------
# Chart 4: Per-point label delete/text and per-point color
#
# officecli add charts-bar.xlsx "/4-Axis & Labels" --type chart \
#   --prop chartType=bar \
#   --prop title="Highlight Winner" \
#   --prop series1="Score:72,85,68,95,78" \
#   --prop categories=Team A,Team B,Team C,Team D,Team E \
#   --prop colors=9DC3E6 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop dataLabels=true --prop labelPos=outsideEnd \
#   --prop dataLabel1.delete=true --prop dataLabel3.delete=true \
#   --prop dataLabel5.delete=true \
#   --prop dataLabel4.text="Winner!" \
#   --prop point4.color=C00000 \
#   --prop point2.color=2E75B6 \
#   --prop gapwidth=70
#
# Features: dataLabel{N}.delete, dataLabel{N}.text, point{N}.color
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Axis & Labels" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Highlight Winner"'
    f' --prop series1=Score:72,85,68,95,78'
    f' --prop categories=Team A,Team B,Team C,Team D,Team E'
    f' --prop colors=9DC3E6'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop dataLabels=true --prop labelPos=outsideEnd'
    f' --prop dataLabel1.delete=true --prop dataLabel3.delete=true'
    f' --prop dataLabel5.delete=true'
    f' --prop dataLabel4.text="Winner!"'
    f' --prop point4.color=C00000'
    f' --prop point2.color=2E75B6'
    f' --prop gapwidth=70')

# ==========================================================================
# Sheet: 5-Legend & Layout
# ==========================================================================
print("\n--- 5-Legend & Layout ---")
cli(f'add "{FILE}" / --type sheet --prop name="5-Legend & Layout"')

# --------------------------------------------------------------------------
# Chart 1: Legend positions (right and bottom)
#
# officecli add charts-bar.xlsx "/5-Legend & Layout" --type chart \
#   --prop chartType=bar \
#   --prop title="Legend: Right" \
#   --prop dataRange=Sheet1!A1:E9 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop colors=4472C4,ED7D31,70AD47,FFC000 \
#   --prop legend=right
#
# Features: legend=right (4-series bar with legend on right)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Legend & Layout" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Legend: Right"'
    f' --prop dataRange=Sheet1!A1:E9'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop colors=4472C4,ED7D31,70AD47,FFC000'
    f' --prop legend=right')

# --------------------------------------------------------------------------
# Chart 2: Legend font styling and overlay
#
# officecli add charts-bar.xlsx "/5-Legend & Layout" --type chart \
#   --prop chartType=bar \
#   --prop title="Legend: Font & Overlay" \
#   --prop dataRange=Sheet1!A1:E9 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop colors=1F4E79,2E75B6,5B9BD5,9DC3E6 \
#   --prop legend=top \
#   --prop legend.overlay=true \
#   --prop legendfont=10:1F4E79:Calibri
#
# Features: legendfont (size:color:fontname), legend.overlay=true
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Legend & Layout" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Legend: Font & Overlay"'
    f' --prop dataRange=Sheet1!A1:E9'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop colors=1F4E79,2E75B6,5B9BD5,9DC3E6'
    f' --prop legend=top'
    f' --prop legend.overlay=true'
    f' --prop legendfont=10:1F4E79:Calibri')

# --------------------------------------------------------------------------
# Chart 3: Manual layout — plotArea.x/y/w/h, title.x/y
#
# officecli add charts-bar.xlsx "/5-Legend & Layout" --type chart \
#   --prop chartType=bar \
#   --prop title="Manual Layout" \
#   --prop dataRange=Sheet1!A1:C9 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop colors=2E75B6,70AD47 \
#   --prop plotArea.x=0.25 --prop plotArea.y=0.15 \
#   --prop plotArea.w=0.70 --prop plotArea.h=0.60 \
#   --prop title.x=0.20 --prop title.y=0.02 \
#   --prop legend.x=0.25 --prop legend.y=0.82 \
#   --prop legend.w=0.50 --prop legend.h=0.10 \
#   --prop title.font=Arial --prop title.size=13 \
#   --prop title.bold=true
#
# Features: plotArea.x/y/w/h, title.x/y, legend.x/y/w/h (manual layout)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Legend & Layout" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Manual Layout"'
    f' --prop dataRange=Sheet1!A1:C9'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop colors=2E75B6,70AD47'
    f' --prop plotArea.x=0.25 --prop plotArea.y=0.15'
    f' --prop plotArea.w=0.70 --prop plotArea.h=0.60'
    f' --prop title.x=0.20 --prop title.y=0.02'
    f' --prop legend.x=0.25 --prop legend.y=0.82'
    f' --prop legend.w=0.50 --prop legend.h=0.10'
    f' --prop title.font=Arial --prop title.size=13'
    f' --prop title.bold=true')

# --------------------------------------------------------------------------
# Chart 4: Secondary axis with chart/plot area borders
#
# officecli add charts-bar.xlsx "/5-Legend & Layout" --type chart \
#   --prop chartType=bar \
#   --prop title="Dual Axis: Revenue vs Margin" \
#   --prop series1="Revenue:340,280,410,195,310" \
#   --prop series2="Margin %:22,18,28,15,25" \
#   --prop categories=North,South,East,West,Central \
#   --prop colors=2E75B6,C00000 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop secondaryAxis=2 \
#   --prop chartArea.border=D0D0D0:1:solid \
#   --prop plotArea.border=E0E0E0:0.5:dot \
#   --prop legend=bottom
#
# Features: secondaryAxis=2, chartArea.border, plotArea.border
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Legend & Layout" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Dual Axis: Revenue vs Margin"'
    f' --prop "series1=Revenue:340,280,410,195,310"'
    f' --prop "series2=Margin %:22,18,28,15,25"'
    f' --prop categories=North,South,East,West,Central'
    f' --prop colors=2E75B6,C00000'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop secondaryAxis=2'
    f' --prop chartArea.border=D0D0D0:1:solid'
    f' --prop plotArea.border=E0E0E0:0.5:dot'
    f' --prop legend=bottom')

# ==========================================================================
# Sheet: 6-Advanced
# ==========================================================================
print("\n--- 6-Advanced ---")
cli(f'add "{FILE}" / --type sheet --prop name="6-Advanced"')

# --------------------------------------------------------------------------
# Chart 1: Reference line with label
#
# officecli add charts-bar.xlsx "/6-Advanced" --type chart \
#   --prop chartType=bar \
#   --prop title="vs Company Average" \
#   --prop series1="Score:82,74,91,68,87,72" \
#   --prop categories=Engineering,Marketing,Sales,Support,Finance,HR \
#   --prop colors=4472C4 \
#   --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
#   --prop referenceLine=79:FF0000:Average:dash \
#   --prop gapwidth=80 \
#   --prop gridlines=E0E0E0:0.5:solid
#
# Features: referenceLine (value:color:label:dash style)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Advanced" --type chart'
    f' --prop chartType=bar'
    f' --prop title="vs Company Average"'
    f' --prop series1=Score:82,74,91,68,87,72'
    f' --prop categories=Engineering,Marketing,Sales,Support,Finance,HR'
    f' --prop colors=4472C4'
    f' --prop x=0 --prop y=0 --prop width=12 --prop height=18'
    f' --prop referenceLine=79:FF0000:Average:dash'
    f' --prop gapwidth=80'
    f' --prop gridlines=E0E0E0:0.5:solid')

# --------------------------------------------------------------------------
# Chart 2: Conditional coloring (colorRule)
#
# officecli add charts-bar.xlsx "/6-Advanced" --type chart \
#   --prop chartType=bar \
#   --prop title="Profit/Loss by Division" \
#   --prop series1="P&L:120,85,-45,160,-80,95,-20,140" \
#   --prop categories=Div A,Div B,Div C,Div D,Div E,Div F,Div G,Div H \
#   --prop colors=2E75B6 \
#   --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
#   --prop colorRule=0:C00000:70AD47 \
#   --prop referenceLine=0:888888:1:solid \
#   --prop dataLabels=outsideEnd \
#   --prop labelFont=9:333333:false
#
# Features: colorRule (threshold:belowColor:aboveColor),
#   referenceLine=0 (zero baseline)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Advanced" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Profit/Loss by Division"'
    f' --prop "series1=P&L:120,85,-45,160,-80,95,-20,140"'
    f' --prop categories=Div A,Div B,Div C,Div D,Div E,Div F,Div G,Div H'
    f' --prop colors=2E75B6'
    f' --prop x=13 --prop y=0 --prop width=12 --prop height=18'
    f' --prop colorRule=0:C00000:70AD47'
    f' --prop referenceLine=0:888888:1:solid'
    f' --prop dataLabels=outsideEnd'
    f' --prop labelFont=9:333333:false')

# --------------------------------------------------------------------------
# Chart 3: Title glow, title shadow, series shadow
#
# officecli add charts-bar.xlsx "/6-Advanced" --type chart \
#   --prop chartType=bar \
#   --prop title="Glow & Shadow Effects" \
#   --prop series1="East:185,195,210,228" \
#   --prop series2="West:140,152,165,178" \
#   --prop categories=Q1,Q2,Q3,Q4 \
#   --prop colors=4472C4,ED7D31 \
#   --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
#   --prop title.glow=4472C4-8-60 \
#   --prop title.shadow=000000-3-315-2-40 \
#   --prop title.font=Calibri --prop title.size=16 \
#   --prop title.bold=true --prop title.color=1F4E79 \
#   --prop series.shadow=000000-3-315-1-30 \
#   --prop plotFill=F0F4F8 --prop chartFill=FFFFFF \
#   --prop legend=bottom
#
# Features: title.glow (color-radius-opacity), title.shadow,
#   series.shadow on bar charts
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Advanced" --type chart'
    f' --prop chartType=bar'
    f' --prop title="Glow & Shadow Effects"'
    f' --prop series1=East:185,195,210,228'
    f' --prop series2=West:140,152,165,178'
    f' --prop categories=Q1,Q2,Q3,Q4'
    f' --prop colors=4472C4,ED7D31'
    f' --prop x=0 --prop y=19 --prop width=12 --prop height=18'
    f' --prop title.glow=4472C4-8-60'
    f' --prop title.shadow=000000-3-315-2-40'
    f' --prop title.font=Calibri --prop title.size=16'
    f' --prop title.bold=true --prop title.color=1F4E79'
    f' --prop series.shadow=000000-3-315-1-30'
    f' --prop plotFill=F0F4F8 --prop chartFill=FFFFFF'
    f' --prop legend=bottom')

# --------------------------------------------------------------------------
# Chart 4: Error bars and data table
#
# officecli add charts-bar.xlsx "/6-Advanced" --type chart \
#   --prop chartType=bar \
#   --prop title="With Error Bars & Data Table" \
#   --prop dataRange=Sheet1!A1:E9 \
#   --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
#   --prop colors=2E75B6,ED7D31,70AD47,FFC000 \
#   --prop errBars=percent:10 \
#   --prop dataTable=true \
#   --prop legend=none \
#   --prop plotFill=FAFAFA
#
# Features: errBars=percent:10, dataTable=true, legend=none
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/6-Advanced" --type chart'
    f' --prop chartType=bar'
    f' --prop title="With Error Bars & Data Table"'
    f' --prop dataRange=Sheet1!A1:E9'
    f' --prop x=13 --prop y=19 --prop width=12 --prop height=18'
    f' --prop colors=2E75B6,ED7D31,70AD47,FFC000'
    f' --prop errBars=percent:10'
    f' --prop dataTable=true'
    f' --prop legend=none'
    f' --prop plotFill=FAFAFA')

print(f"\nDone! Generated: {FILE}")
print("  7 sheets (Sheet1 data + 6 chart sheets, 24 charts total)")
