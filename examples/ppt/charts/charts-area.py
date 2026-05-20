#!/usr/bin/env python3
"""
Area Charts Showcase — area, stackedArea, percentStackedArea, area3d.

Generates: charts-area.pptx

  Slide 1  Variants           area / stackedArea / percentStackedArea / area3d
  Slide 2  Title & legend     title.* + legend positions + legendFont
  Slide 3  Data labels        flags + labelPos + labelfont
  Slide 4  Axes               min/max, titles, fonts, gridlines, ticks, labelrotation
  Slide 5  Series styling     colors, gradient, gradients, transparency, seriesoutline, seriesshadow
  Slide 6  Overlays           referenceline, errbars, trendline
  Slide 7  Backgrounds        chartareafill, plotFill, chartborder, plotborder, roundedcorners
  Slide 8  Presets & per-ser  preset bundles + seriesN.* + chart-series Set

Usage:
  python3 charts-area.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-area.pptx")
def cli(*a):
    r = subprocess.run(["officecli", *a], capture_output=True, text=True)
    if r.returncode:
        m=(r.stderr or r.stdout or "").strip().splitlines()
        head=m[0][:160] if m else ""
        if "UNSUPPORTED" in (r.stderr or ""):
            # Forward-compat: skip unsupported props but surface so silent gaps are visible.
            print(f"  ⚠ {' '.join(a[:3])} → {head}", file=sys.stderr); return
        if m: print(f"  ! {' '.join(a[:3])} → {head}", file=sys.stderr)
        sys.exit(r.returncode)
def P(d): return [x for k,v in d.items() for x in ("--prop", f"{k}={v}")]
slide=0
def new_slide(t):
    global slide; slide+=1
    cli("add",FILE,"/","--type","slide")
    cli("add",FILE,f"/slide[{slide}]","--type","shape",
        *P({"text":t,"size":24,"bold":"true","autoFit":"normal","x":"0.5in","y":"0.3in","width":"12.3in","height":"0.6in"}))
def ch(box,p): cli("add",FILE,f"/slide[{slide}]","--type","chart",*P({**box,**p}))
TL={"x":"0.3in","y":"1.05in","width":"6.1in","height":"3in"}
TR={"x":"6.95in","y":"1.05in","width":"6.1in","height":"3in"}
BL={"x":"0.3in","y":"4.25in","width":"6.1in","height":"3in"}
BR={"x":"6.95in","y":"4.25in","width":"6.1in","height":"3in"}
CATS="Mon,Tue,Wed,Thu,Fri"
D="A:50,60,70,65,80"
D2="Web:50,60,70,65,80;Mobile:30,35,42,48,55"

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

new_slide("Area variants — area / stackedArea / percentStackedArea / area3d")
ch(TL,{"chartType":"area","title":"area","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"stackedArea","title":"stackedArea","legend":"bottom","categories":CATS,"data":D2})
ch(BL,{"chartType":"percentStackedArea","title":"percentStackedArea","legend":"bottom","categories":CATS,"data":D2})
ch(BR,{"chartType":"area3d","title":"area3d","view3d":"15,20,30","legend":"bottom","categories":CATS,"data":D2})

new_slide("Title & legend")
ch(TL,{"chartType":"area","title":"Styled title","title.font":"Georgia","title.size":"20",
       "title.color":"4472C4","title.bold":"true","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"area","title":"legend=top + legendFont","legend":"top",
       "legendFont":"10:333333:Calibri","categories":CATS,"data":D2})
ch(BL,{"chartType":"area","title":"legend.overlay=true","legend":"topRight",
       "legend.overlay":"true","categories":CATS,"data":D2})
ch(BR,{"chartType":"area","autotitledeleted":"true","legend":"none","categories":CATS,"data":D2})

new_slide("Data labels — flags, labelPos, labelfont")
ch(TL,{"chartType":"area","title":"dataLabels=value","dataLabels":"value",
       "labelfont":"10:333333:Calibri","legend":"none","categories":CATS,"data":D})
ch(TR,{"chartType":"stackedArea","title":"stacked + center labels","dataLabels":"value",
       "labelPos":"center","legend":"bottom","categories":CATS,"data":D2})
ch(BL,{"chartType":"area","title":"value,category","dataLabels":"value,category",
       "labelfont":"9:333333:Calibri","legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"area","title":"dataLabels=none","dataLabels":"none","legend":"none",
       "categories":CATS,"data":D})

new_slide("Axes — min/max, gridlines, ticks, labelrotation")
ch(TL,{"chartType":"area","title":"min/max + titles","legend":"none",
       "axismin":"0","axismax":"100","majorunit":"25","axistitle":"Value","cattitle":"Day",
       "axisfont":"10:333333:Calibri","axisline":"666666:1","axisnumfmt":"#,##0",
       "categories":CATS,"data":D})
ch(TR,{"chartType":"area","title":"gridlines + ticks","legend":"none",
       "gridlines":"E0E0E0:0.3","minorGridlines":"F0F0F0:0.25",
       "majorTickMark":"out","minorTickMark":"in","tickLabelPos":"nextTo",
       "categories":CATS,"data":D})
ch(BL,{"chartType":"area","title":"labelrotation=-30","legend":"none","labelrotation":"-30",
       "categories":"January,February,March,April,May,June","data":"A:60,90,140,180,160,210"})
ch(BR,{"chartType":"area","title":"dispunits=thousands","legend":"none","dispunits":"thousands",
       "categories":CATS,"data":"Rev:120000,135000,148000,162000,180000"})

new_slide("Series styling — colors, gradient(s), transparency, outline, shadow")
ch(TL,{"chartType":"area","title":"colors + seriesoutline","legend":"bottom",
       "colors":"4472C4,ED7D31","seriesoutline":"000000:0.5","categories":CATS,"data":D2})
ch(TR,{"chartType":"area","title":"gradient + seriesshadow","legend":"none",
       "gradient":"FF6600-FFCC00:90","seriesshadow":"000000-5-45-3-50",
       "categories":CATS,"data":D})
ch(BL,{"chartType":"area","title":"per-series gradients + transparency=30",
       "gradients":"FF0000-0000FF;00FF00-FFFF00","transparency":"30",
       "legend":"bottom","categories":CATS,"data":D2})
ch(BR,{"chartType":"area","title":"single + transparency=50","transparency":"50",
       "colors":"4472C4","legend":"none","categories":CATS,"data":D})

new_slide("Overlays — referenceline, errbars, trendline")
ch(TL,{"chartType":"area","title":"referenceline=60","referenceline":"60:FF0000:Target",
       "legend":"none","categories":CATS,"data":D})
ch(TR,{"chartType":"area","title":"errbars=percentage:10","errbars":"percentage:10",
       "legend":"none","categories":CATS,"data":D})
ch(BL,{"chartType":"area","title":"trendline=linear","trendline":"linear",
       "legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"area","title":"trendline=movingAvg:3","trendline":"movingAvg:3",
       "legend":"none","categories":CATS,"data":D})

new_slide("Backgrounds — chartareafill, plotFill, chartborder, plotborder, roundedcorners")
ch(TL,{"chartType":"area","title":"chartareafill + plotFill + borders","legend":"bottom",
       "chartareafill":"FFF8E7","plotFill":"FAFAFA","chartborder":"000000:1",
       "plotborder":"CCCCCC:0.5","categories":CATS,"data":D2})
ch(TR,{"chartType":"area","title":"roundedcorners=true","roundedcorners":"true",
       "chartborder":"4472C4:2","legend":"bottom","categories":CATS,"data":D2})
ch(BL,{"chartType":"area","title":"plotFill=none","plotFill":"none","gridlines":"none",
       "legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"area","title":"dataTable=true","dataTable":"true","legend":"bottom",
       "categories":CATS,"data":D2})

new_slide("Presets & per-series control")
for box,p in zip([TL,TR,BL],["minimal","dark","corporate"]):
    ch(box,{"chartType":"area","preset":p,"title":f"preset={p}","legend":"bottom",
            "categories":CATS,"data":D2})
ch(BR,{"chartType":"area","title":"seriesN.* + chart-series Set","legend":"bottom",
       "categories":CATS,
       "series1.name":"Web","series1.values":"50,60,70,65,80","series1.color":"4472C4",
       "series2.name":"Mobile","series2.values":"30,35,42,48,55","series2.color":"ED7D31"})
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[1]",*P({"name":"Renamed Web","color":"C00000"}))

print(f"Done: {FILE}  ({slide} slides)")
