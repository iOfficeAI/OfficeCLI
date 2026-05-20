#!/usr/bin/env python3
"""
Scatter Charts Showcase — scatterstyle line/lineMarker/marker/smooth/smoothMarker.

Generates: charts-scatter.pptx

  Slide 1  scatterstyle variants  line / lineMarker / marker / smooth / smoothMarker (5 charts)
  Slide 2  Markers                marker symbol/size/color
  Slide 3  Title & legend
  Slide 4  Data labels
  Slide 5  Axes                   min/max, gridlines, log on both axes
  Slide 6  Series styling         colors, gradient, transparency, outline, shadow
  Slide 7  Overlays               trendline (linear/poly/exp/log/power/movingAvg), errbars, referenceline
  Slide 8  Per-series Set         lineWidth/lineDash/marker/markerSize/color/smooth + presets

Usage:
  python3 charts-scatter.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-scatter.pptx")
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
D="A:10,20,18,30,28,40,42,55,52,65"
D2="A:10,20,18,30,28,40,42,55;B:5,12,15,22,25,30,35,40"

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

new_slide("scatterstyle — line / lineMarker / marker / smooth / smoothMarker")
ch(TL,{"chartType":"scatter","scatterstyle":"line","title":"scatterstyle=line",
       "legend":"none","data":D})
ch(TR,{"chartType":"scatter","scatterstyle":"lineMarker","title":"scatterstyle=lineMarker",
       "legend":"none","data":D})
ch(BL,{"chartType":"scatter","scatterstyle":"marker","title":"scatterstyle=marker",
       "legend":"none","data":D})
ch(BR,{"chartType":"scatter","scatterstyle":"smoothMarker","title":"scatterstyle=smoothMarker",
       "legend":"none","data":D})

new_slide("Markers — symbol / size / color")
ch(TL,{"chartType":"scatter","scatterstyle":"marker","title":"circle:10:FF0000",
       "marker":"circle:10:FF0000","legend":"none","data":D})
ch(TR,{"chartType":"scatter","scatterstyle":"marker","title":"diamond:12:0070C0",
       "marker":"diamond:12:0070C0","legend":"none","data":D})
ch(BL,{"chartType":"scatter","scatterstyle":"marker","title":"square:8:70AD47",
       "marker":"square:8:70AD47","legend":"none","data":D})
ch(BR,{"chartType":"scatter","scatterstyle":"marker","title":"triangle:10",
       "marker":"triangle:10","legend":"none","data":D})

new_slide("Title & legend")
ch(TL,{"chartType":"scatter","scatterstyle":"smoothMarker","title":"Styled title",
       "title.font":"Georgia","title.size":"20","title.color":"4472C4","title.bold":"true",
       "legend":"bottom","data":D2})
ch(TR,{"chartType":"scatter","scatterstyle":"lineMarker","title":"legend=top + legendFont",
       "legend":"top","legendFont":"10:333333:Calibri","data":D2})
ch(BL,{"chartType":"scatter","scatterstyle":"lineMarker","title":"legend.overlay=true",
       "legend":"topRight","legend.overlay":"true","data":D2})
ch(BR,{"chartType":"scatter","scatterstyle":"marker","autotitledeleted":"true","legend":"none","data":D2})

new_slide("Data labels — flags + labelfont")
ch(TL,{"chartType":"scatter","scatterstyle":"marker","title":"value","dataLabels":"value",
       "labelfont":"9:333333:Calibri","legend":"none","data":D})
ch(TR,{"chartType":"scatter","scatterstyle":"marker","title":"value,series",
       "dataLabels":"value,series","legend":"none","data":D2})
ch(BL,{"chartType":"scatter","scatterstyle":"marker","title":"labelPos=top",
       "dataLabels":"value","labelPos":"top","legend":"none","data":D})
ch(BR,{"chartType":"scatter","scatterstyle":"marker","title":"dataLabels=none",
       "dataLabels":"none","legend":"none","data":D})

new_slide("Axes — min/max, gridlines, ticks, log on both axes")
ch(TL,{"chartType":"scatter","scatterstyle":"lineMarker","title":"min/max + titles",
       "axismin":"0","axismax":"80","majorunit":"20","axistitle":"Y","cattitle":"X",
       "axisfont":"10:333333:Calibri","axisline":"666666:1","axisnumfmt":"#,##0",
       "legend":"none","data":D})
ch(TR,{"chartType":"scatter","scatterstyle":"marker","title":"gridlines + minorGridlines",
       "gridlines":"E0E0E0:0.3","minorGridlines":"F0F0F0:0.25","legend":"none","data":D})
ch(BL,{"chartType":"scatter","scatterstyle":"marker","title":"labelrotation=-30",
       "labelrotation":"-30","legend":"none","data":D})
ch(BR,{"chartType":"scatter","scatterstyle":"marker","title":"logbase=10 (Y)",
       "logbase":"10","axismin":"1","axismax":"100","legend":"none",
       "data":"A:2,5,8,12,20,40,80"})

new_slide("Series styling — colors, gradient, transparency, outline, shadow")
ch(TL,{"chartType":"scatter","scatterstyle":"marker","title":"colors + seriesoutline",
       "colors":"4472C4,ED7D31","seriesoutline":"000000:0.5","legend":"bottom","data":D2})
ch(TR,{"chartType":"scatter","scatterstyle":"marker","title":"gradient + seriesshadow",
       "gradient":"FF6600-FFCC00","seriesshadow":"000000-5-45-3-50","legend":"none","data":D})
ch(BL,{"chartType":"scatter","scatterstyle":"marker","title":"transparency=30",
       "transparency":"30","legend":"bottom","data":D2})
ch(BR,{"chartType":"scatter","scatterstyle":"marker","title":"per-series gradients",
       "gradients":"FF0000-0000FF;00FF00-FFFF00","legend":"bottom","data":D2})

new_slide("Overlays — trendline (linear/poly/exp/movingAvg), errbars, referenceline")
ch(TL,{"chartType":"scatter","scatterstyle":"marker","title":"trendline=linear",
       "trendline":"linear","legend":"none","data":D})
ch(TR,{"chartType":"scatter","scatterstyle":"marker","title":"trendline=poly:3",
       "trendline":"poly:3","legend":"none","data":D})
ch(BL,{"chartType":"scatter","scatterstyle":"marker","title":"trendline=movingAvg:3",
       "trendline":"movingAvg:3","legend":"none","data":D})
ch(BR,{"chartType":"scatter","scatterstyle":"marker","title":"errbars=stdDev:1",
       "errbars":"stdDev:1","legend":"none","data":D})

new_slide("Per-series Set + presets — lineWidth/lineDash/marker/markerSize/color/smooth")
for box,p in zip([TL,TR,BL],["minimal","dark","corporate"]):
    ch(box,{"chartType":"scatter","scatterstyle":"smoothMarker","preset":p,
            "title":f"preset={p}","legend":"bottom","data":D2})
ch(BR,{"chartType":"scatter","scatterstyle":"lineMarker","title":"chart-series Set per series",
       "legend":"bottom","data":D2})
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[1]",
    *P({"name":"Alpha","color":"C00000","lineWidth":"2.5","lineDash":"solid",
        "marker":"circle","markerSize":"10","smooth":"true"}))
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[2]",
    *P({"name":"Beta","color":"2E75B6","lineWidth":"1.5","lineDash":"dash",
        "marker":"diamond","markerSize":"8"}))

print(f"Done: {FILE}  ({slide} slides)")
