#!/usr/bin/env python3
"""
Pie Charts Showcase — pie, pie3d, pieOfPie, barOfPie (where supported).

Generates: charts-pie.pptx

  Slide 1  Variants           pie / pie3d (view3d) — varyColors, firstSliceAngle
  Slide 2  Explosion          explosion=0/10/20/30
  Slide 3  Title & legend     title.* + legend positions + legendFont
  Slide 4  Data labels        flags (percent/category/value), labelfont, leaderlines
  Slide 5  Series styling     colors, gradient, transparency, seriesoutline, seriesshadow
  Slide 6  First-slice angle  0 / 90 / 180 / 270
  Slide 7  Backgrounds        chartareafill, plotFill, chartborder, roundedcorners
  Slide 8  Presets & per-pt   preset bundles + per-point recolor via chart-series Set

Usage:
  python3 charts-pie.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-pie.pptx")
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

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

new_slide("Pie variants — pie / pie3d (varyColors, firstSliceAngle)")
ch(TL,{"chartType":"pie","title":"pie","legend":"right","varyColors":"true",
       "categories":CATS,"data":D})
ch(TR,{"chartType":"pie3d","title":"pie3d (view3d=20,20,30)","view3d":"20,20,30",
       "legend":"right","varyColors":"true","categories":CATS,"data":D})
ch(BL,{"chartType":"pie","title":"firstSliceAngle=90","firstSliceAngle":"90",
       "legend":"right","categories":CATS,"data":D})
ch(BR,{"chartType":"pie","title":"varyColors=false","varyColors":"false",
       "legend":"right","categories":CATS,"data":D})

new_slide("Explosion — 0 / 10 / 20 / 30 (% of radius)")
ch(TL,{"chartType":"pie","title":"explosion=0","explosion":"0","legend":"right",
       "categories":CATS,"data":D})
ch(TR,{"chartType":"pie","title":"explosion=10","explosion":"10","legend":"right",
       "categories":CATS,"data":D})
ch(BL,{"chartType":"pie","title":"explosion=20","explosion":"20","legend":"right",
       "categories":CATS,"data":D})
ch(BR,{"chartType":"pie","title":"explosion=30","explosion":"30","legend":"right",
       "categories":CATS,"data":D})

new_slide("Title & legend")
ch(TL,{"chartType":"pie","title":"Styled title","title.font":"Georgia","title.size":"20",
       "title.color":"4472C4","title.bold":"true","legend":"right","categories":CATS,"data":D})
ch(TR,{"chartType":"pie","title":"legend=bottom + legendFont","legend":"bottom",
       "legendFont":"10:333333:Calibri","categories":CATS,"data":D})
ch(BL,{"chartType":"pie","title":"legend.overlay=true","legend":"topRight",
       "legend.overlay":"true","categories":CATS,"data":D})
ch(BR,{"chartType":"pie","autotitledeleted":"true","legend":"none","categories":CATS,"data":D})

new_slide("Data labels — percent / category / value, labelfont, leaderlines")
ch(TL,{"chartType":"pie","title":"dataLabels=percent","dataLabels":"percent",
       "legend":"right","labelfont":"10:333333:Calibri","categories":CATS,"data":D})
ch(TR,{"chartType":"pie","title":"percent,category","dataLabels":"percent,category",
       "leaderlines":"true","legend":"none","labelfont":"10:333333:Calibri",
       "categories":CATS,"data":D})
ch(BL,{"chartType":"pie","title":"all flags","dataLabels":"value,percent,category",
       "leaderlines":"true","legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"pie","title":"dataLabels=none","dataLabels":"none","legend":"right",
       "categories":CATS,"data":D})

new_slide("Series styling — colors, gradient, transparency, outline, shadow")
ch(TL,{"chartType":"pie","title":"colors= explicit palette","legend":"right",
       "colors":"4472C4,ED7D31,A5A5A5,70AD47","categories":CATS,"data":D})
ch(TR,{"chartType":"pie","title":"gradient + seriesshadow","legend":"right",
       "gradient":"FF6600-FFCC00","seriesshadow":"000000-5-45-3-50",
       "categories":CATS,"data":D})
ch(BL,{"chartType":"pie","title":"seriesoutline white","legend":"right",
       "seriesoutline":"FFFFFF:2","categories":CATS,"data":D})
ch(BR,{"chartType":"pie","title":"transparency=30","legend":"right",
       "transparency":"30","categories":CATS,"data":D})

new_slide("First slice angle — 0 / 90 / 180 / 270")
for box,ang in zip([TL,TR,BL,BR],[0,90,180,270]):
    ch(box,{"chartType":"pie","title":f"firstSliceAngle={ang}",
            "firstSliceAngle":str(ang),"legend":"right",
            "varyColors":"true","categories":CATS,"data":D})

new_slide("Backgrounds — chartareafill, plotFill, chartborder, roundedcorners")
ch(TL,{"chartType":"pie","title":"chartareafill + chartborder","legend":"right",
       "chartareafill":"FFF8E7","chartborder":"000000:1","categories":CATS,"data":D})
ch(TR,{"chartType":"pie","title":"roundedcorners=true","legend":"right",
       "roundedcorners":"true","chartborder":"4472C4:2","categories":CATS,"data":D})
ch(BL,{"chartType":"pie","title":"plotFill=none","legend":"right",
       "plotFill":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"pie","title":"chartareafill=none","legend":"right",
       "chartareafill":"none","categories":CATS,"data":D})

new_slide("Presets & per-series Set")
for box,p in zip([TL,TR,BL],["minimal","dark","corporate"]):
    ch(box,{"chartType":"pie","preset":p,"title":f"preset={p}","legend":"right",
            "categories":CATS,"data":D})
ch(BR,{"chartType":"pie","title":"chart-series Set name+color","legend":"right",
       "categories":CATS,"data":D})
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[1]",
    *P({"name":"Renamed Share","color":"C00000"}))

print(f"Done: {FILE}  ({slide} slides)")
