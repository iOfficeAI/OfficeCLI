#!/usr/bin/env python3
"""
Doughnut Charts Showcase.

Generates: charts-doughnut.pptx

  Slide 1  holeSize variants      holeSize=10/30/55/75
  Slide 2  Multi-ring             two-series + three-series concentric rings
  Slide 3  firstSliceAngle        0 / 90 / 180 / 270
  Slide 4  Data labels            percent / category / value, leaderlines, labelfont
  Slide 5  Series styling         colors, gradient, seriesoutline, seriesshadow, transparency
  Slide 6  Title & legend         title.* + legend positions + legendFont
  Slide 7  Backgrounds            chartareafill, plotFill, chartborder, roundedcorners
  Slide 8  Presets & per-series   preset bundles + chart-series Set

Usage:
  python3 charts-doughnut.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-doughnut.pptx")
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
CATS="North,South,East,West"; D="Share:30,25,28,17"
D2="Last:25,30,25,20;This:30,25,28,17"
D3="Region1:30,25,28,17;Region2:25,30,20,25;Region3:20,25,30,25"

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

new_slide("holeSize — 10 / 30 / 55 / 75")
for box,h in zip([TL,TR,BL,BR],[10,30,55,75]):
    ch(box,{"chartType":"doughnut","title":f"holeSize={h}","holeSize":str(h),
            "legend":"right","varyColors":"true","categories":CATS,"data":D})

new_slide("Multi-ring — concentric series")
ch(TL,{"chartType":"doughnut","title":"single ring","holeSize":"50","legend":"right",
       "categories":CATS,"data":D})
ch(TR,{"chartType":"doughnut","title":"two rings","holeSize":"40","legend":"right",
       "categories":CATS,"data":D2})
ch(BL,{"chartType":"doughnut","title":"three rings","holeSize":"30","legend":"right",
       "categories":CATS,"data":D3})
ch(BR,{"chartType":"doughnut","title":"two rings + dataLabels=percent","holeSize":"40",
       "dataLabels":"percent","legend":"right","categories":CATS,"data":D2})

new_slide("First slice angle — 0 / 90 / 180 / 270")
for box,ang in zip([TL,TR,BL,BR],[0,90,180,270]):
    ch(box,{"chartType":"doughnut","title":f"firstSliceAngle={ang}",
            "firstSliceAngle":str(ang),"holeSize":"50","legend":"right",
            "varyColors":"true","categories":CATS,"data":D})

new_slide("Data labels — percent / category / value, leaderlines, labelfont")
ch(TL,{"chartType":"doughnut","title":"dataLabels=percent","dataLabels":"percent","holeSize":"50",
       "legend":"right","labelfont":"10:333333:Calibri","categories":CATS,"data":D})
ch(TR,{"chartType":"doughnut","title":"percent,category","dataLabels":"percent,category",
       "holeSize":"50","leaderlines":"true","legend":"none",
       "labelfont":"10:333333:Calibri","categories":CATS,"data":D})
ch(BL,{"chartType":"doughnut","title":"all flags","dataLabels":"value,percent,category",
       "holeSize":"50","leaderlines":"true","legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"doughnut","title":"dataLabels=none","dataLabels":"none","holeSize":"50",
       "legend":"right","categories":CATS,"data":D})

new_slide("Series styling — colors, gradient, outline, shadow, transparency")
ch(TL,{"chartType":"doughnut","title":"colors=","holeSize":"50","legend":"right",
       "colors":"4472C4,ED7D31,A5A5A5,70AD47","categories":CATS,"data":D})
ch(TR,{"chartType":"doughnut","title":"gradient + seriesshadow","holeSize":"50",
       "gradient":"FF6600-FFCC00","seriesshadow":"000000-5-45-3-50",
       "legend":"right","categories":CATS,"data":D})
ch(BL,{"chartType":"doughnut","title":"seriesoutline white","holeSize":"50",
       "seriesoutline":"FFFFFF:2","legend":"right","categories":CATS,"data":D})
ch(BR,{"chartType":"doughnut","title":"transparency=30","holeSize":"50",
       "transparency":"30","legend":"right","categories":CATS,"data":D})

new_slide("Title & legend")
ch(TL,{"chartType":"doughnut","title":"Styled title","title.font":"Georgia","title.size":"20",
       "title.color":"4472C4","title.bold":"true","holeSize":"50","legend":"right",
       "categories":CATS,"data":D})
ch(TR,{"chartType":"doughnut","title":"legend=bottom + legendFont","holeSize":"50",
       "legend":"bottom","legendFont":"10:333333:Calibri","categories":CATS,"data":D})
ch(BL,{"chartType":"doughnut","title":"legend.overlay=true","holeSize":"50",
       "legend":"topRight","legend.overlay":"true","categories":CATS,"data":D})
ch(BR,{"chartType":"doughnut","autotitledeleted":"true","holeSize":"50","legend":"none",
       "categories":CATS,"data":D})

new_slide("Backgrounds — chartareafill, plotFill, chartborder, roundedcorners")
ch(TL,{"chartType":"doughnut","title":"chartareafill + chartborder","holeSize":"50",
       "chartareafill":"FFF8E7","chartborder":"000000:1","legend":"right",
       "categories":CATS,"data":D})
ch(TR,{"chartType":"doughnut","title":"roundedcorners=true","holeSize":"50",
       "roundedcorners":"true","chartborder":"4472C4:2","legend":"right",
       "categories":CATS,"data":D})
ch(BL,{"chartType":"doughnut","title":"plotFill=none","holeSize":"50",
       "plotFill":"none","legend":"right","categories":CATS,"data":D})
ch(BR,{"chartType":"doughnut","title":"chartareafill=none","holeSize":"50",
       "chartareafill":"none","legend":"right","categories":CATS,"data":D})

new_slide("Presets & per-series Set")
for box,p in zip([TL,TR,BL],["minimal","dark","corporate"]):
    ch(box,{"chartType":"doughnut","preset":p,"title":f"preset={p}","holeSize":"50",
            "legend":"right","categories":CATS,"data":D})
ch(BR,{"chartType":"doughnut","title":"chart-series Set name+color","holeSize":"50",
       "legend":"right","categories":CATS,"data":D})
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[1]",
    *P({"name":"Renamed Share","color":"C00000"}))

print(f"Done: {FILE}  ({slide} slides)")
