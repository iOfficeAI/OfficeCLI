#!/usr/bin/env python3
"""
Radar Charts Showcase — radarstyle standard / marker / filled.

Generates: charts-radar.pptx

  Slide 1  radarstyle             standard / marker / filled
  Slide 2  Title & legend         title.* + legend positions + legendFont
  Slide 3  Data labels            flags + labelfont
  Slide 4  Axes                   min/max, gridlines, axisfont, labelrotation
  Slide 5  Series styling         colors, gradient, transparency, outline, shadow
  Slide 6  Markers                marker symbol/size/color (radarstyle=marker only)
  Slide 7  Backgrounds            chartareafill, plotFill, chartborder, roundedcorners
  Slide 8  Presets & per-series   preset bundles + chart-series Set

Usage:
  python3 charts-radar.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-radar.pptx")
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
CATS="Speed,Power,Range,Style,Tech,Price"
D="A:8,7,9,6,8,7"
D2="Model A:8,7,9,6,8,7;Model B:6,9,7,8,9,6"

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

new_slide("radarstyle — standard / marker / filled")
ch(TL,{"chartType":"radar","radarstyle":"standard","title":"radarstyle=standard",
       "legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"radar","radarstyle":"marker","title":"radarstyle=marker",
       "legend":"bottom","categories":CATS,"data":D2})
ch(BL,{"chartType":"radar","radarstyle":"filled","title":"radarstyle=filled",
       "legend":"bottom","categories":CATS,"data":D2})
ch(BR,{"chartType":"radar","radarstyle":"standard","title":"single series",
       "legend":"bottom","categories":CATS,"data":D})

new_slide("Title & legend")
ch(TL,{"chartType":"radar","radarstyle":"filled","title":"Styled title",
       "title.font":"Georgia","title.size":"20","title.color":"4472C4","title.bold":"true",
       "legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"radar","radarstyle":"standard","title":"legend=top + legendFont",
       "legend":"top","legendFont":"10:333333:Calibri","categories":CATS,"data":D2})
ch(BL,{"chartType":"radar","radarstyle":"standard","title":"legend.overlay=true",
       "legend":"topRight","legend.overlay":"true","categories":CATS,"data":D2})
ch(BR,{"chartType":"radar","radarstyle":"filled","autotitledeleted":"true","legend":"none",
       "categories":CATS,"data":D2})

new_slide("Data labels — flags + labelfont")
ch(TL,{"chartType":"radar","radarstyle":"marker","title":"value","dataLabels":"value",
       "labelfont":"9:333333:Calibri","legend":"none","categories":CATS,"data":D})
ch(TR,{"chartType":"radar","radarstyle":"marker","title":"value,series",
       "dataLabels":"value,series","legend":"bottom","categories":CATS,"data":D2})
ch(BL,{"chartType":"radar","radarstyle":"standard","title":"value,category",
       "dataLabels":"value,category","legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"radar","radarstyle":"filled","title":"dataLabels=none",
       "dataLabels":"none","legend":"bottom","categories":CATS,"data":D2})

new_slide("Axes — min/max, gridlines, axisfont, labelrotation")
ch(TL,{"chartType":"radar","radarstyle":"standard","title":"min/max + titles",
       "axismin":"0","axismax":"10","majorunit":"2","axisfont":"10:333333:Calibri",
       "legend":"none","categories":CATS,"data":D})
ch(TR,{"chartType":"radar","radarstyle":"standard","title":"gridlines + minorGridlines",
       "gridlines":"E0E0E0:0.3","minorGridlines":"F0F0F0:0.25",
       "legend":"none","categories":CATS,"data":D})
ch(BL,{"chartType":"radar","radarstyle":"standard","title":"labelrotation=30",
       "labelrotation":"30","legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"radar","radarstyle":"standard","title":"axisnumfmt=0.0",
       "axisnumfmt":"0.0","legend":"none","categories":CATS,"data":D})

new_slide("Series styling — colors, gradient, transparency, outline, shadow")
ch(TL,{"chartType":"radar","radarstyle":"filled","title":"colors + seriesoutline",
       "colors":"4472C4,ED7D31","seriesoutline":"000000:0.5",
       "legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"radar","radarstyle":"filled","title":"gradient + seriesshadow",
       "gradient":"FF6600-FFCC00","seriesshadow":"000000-5-45-3-50",
       "legend":"none","categories":CATS,"data":D})
ch(BL,{"chartType":"radar","radarstyle":"filled","title":"transparency=40",
       "transparency":"40","legend":"bottom","categories":CATS,"data":D2})
ch(BR,{"chartType":"radar","radarstyle":"filled","title":"per-series gradients",
       "gradients":"FF0000-0000FF;00FF00-FFFF00","legend":"bottom","categories":CATS,"data":D2})

new_slide("Markers (radarstyle=marker) — symbol/size/color")
ch(TL,{"chartType":"radar","radarstyle":"marker","title":"circle:10:FF0000",
       "marker":"circle:10:FF0000","legend":"none","categories":CATS,"data":D})
ch(TR,{"chartType":"radar","radarstyle":"marker","title":"square:8:0070C0",
       "marker":"square:8:0070C0","legend":"none","categories":CATS,"data":D})
ch(BL,{"chartType":"radar","radarstyle":"marker","title":"diamond:12",
       "marker":"diamond:12","legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"radar","radarstyle":"marker","title":"triangle:10:70AD47",
       "marker":"triangle:10:70AD47","legend":"none","categories":CATS,"data":D})

new_slide("Backgrounds — chartareafill, plotFill, chartborder, roundedcorners")
ch(TL,{"chartType":"radar","radarstyle":"filled","title":"chartareafill + plotFill + borders",
       "chartareafill":"FFF8E7","plotFill":"FAFAFA","chartborder":"000000:1",
       "plotborder":"CCCCCC:0.5","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"radar","radarstyle":"filled","title":"roundedcorners=true",
       "roundedcorners":"true","chartborder":"4472C4:2",
       "legend":"bottom","categories":CATS,"data":D2})
ch(BL,{"chartType":"radar","radarstyle":"standard","title":"plotFill=none",
       "plotFill":"none","legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"radar","radarstyle":"filled","title":"chartareafill=none",
       "chartareafill":"none","legend":"bottom","categories":CATS,"data":D2})

new_slide("Presets & per-series Set")
for box,p in zip([TL,TR,BL],["minimal","dark","corporate"]):
    ch(box,{"chartType":"radar","radarstyle":"filled","preset":p,"title":f"preset={p}",
            "legend":"bottom","categories":CATS,"data":D2})
ch(BR,{"chartType":"radar","radarstyle":"marker","title":"chart-series Set",
       "legend":"bottom","categories":CATS,"data":D2})
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[1]",
    *P({"name":"Renamed A","color":"C00000","marker":"circle","markerSize":"9"}))

print(f"Done: {FILE}  ({slide} slides)")
