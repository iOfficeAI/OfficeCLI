#!/usr/bin/env python3
"""
Bubble Charts Showcase.

Generates: charts-bubble.pptx

  Slide 1  bubbleScale            50 / 100 / 150 / 200 (% of default)
  Slide 2  sizerepresents         area vs width
  Slide 3  shownegbubbles         true vs false (with negative values)
  Slide 4  Title & legend         title.* + legend positions + legendFont
  Slide 5  Data labels            value/category/bubbleSize, labelfont
  Slide 6  Axes                   min/max, gridlines, ticks
  Slide 7  Series styling         colors, gradient, transparency, outline, shadow
  Slide 8  Presets & per-series   preset bundles + chart-series Set

Usage:
  python3 charts-bubble.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-bubble.pptx")
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
D="A:5,12,8,18,22,9,15,11"
D2="A:5,12,8,18,22,9;B:7,11,15,9,20,14"

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

new_slide("bubbleScale — 50 / 100 / 150 / 200 (% of default)")
for box,s in zip([TL,TR,BL,BR],[50,100,150,200]):
    ch(box,{"chartType":"bubble","title":f"bubbleScale={s}","bubbleScale":str(s),
            "legend":"none","data":D})

new_slide("sizerepresents — area vs width")
ch(TL,{"chartType":"bubble","title":"sizerepresents=area","sizerepresents":"area",
       "legend":"none","data":D})
ch(TR,{"chartType":"bubble","title":"sizerepresents=width","sizerepresents":"width",
       "legend":"none","data":D})
ch(BL,{"chartType":"bubble","title":"area + 2 series","sizerepresents":"area",
       "legend":"bottom","data":D2})
ch(BR,{"chartType":"bubble","title":"width + 2 series","sizerepresents":"width",
       "legend":"bottom","data":D2})

new_slide("shownegbubbles — false vs true")
ch(TL,{"chartType":"bubble","title":"shownegbubbles=false","shownegbubbles":"false",
       "legend":"none","data":"A:5,-8,12,-15,18,22"})
ch(TR,{"chartType":"bubble","title":"shownegbubbles=true","shownegbubbles":"true",
       "legend":"none","data":"A:5,-8,12,-15,18,22"})
ch(BL,{"chartType":"bubble","title":"false + 2 series","shownegbubbles":"false",
       "legend":"bottom","data":"A:5,-8,12,-15,18,22;B:8,11,-9,14,-16,20"})
ch(BR,{"chartType":"bubble","title":"true + 2 series","shownegbubbles":"true",
       "legend":"bottom","data":"A:5,-8,12,-15,18,22;B:8,11,-9,14,-16,20"})

new_slide("Title & legend")
ch(TL,{"chartType":"bubble","title":"Styled title","title.font":"Georgia","title.size":"20",
       "title.color":"4472C4","title.bold":"true","legend":"bottom","data":D2})
ch(TR,{"chartType":"bubble","title":"legend=top + legendFont","legend":"top",
       "legendFont":"10:333333:Calibri","data":D2})
ch(BL,{"chartType":"bubble","title":"legend.overlay=true","legend":"topRight",
       "legend.overlay":"true","data":D2})
ch(BR,{"chartType":"bubble","autotitledeleted":"true","legend":"none","data":D2})

new_slide("Data labels — flags + labelfont")
ch(TL,{"chartType":"bubble","title":"value","dataLabels":"value",
       "labelfont":"9:333333:Calibri","legend":"none","data":D})
ch(TR,{"chartType":"bubble","title":"value,series","dataLabels":"value,series",
       "legend":"none","data":D2})
ch(BL,{"chartType":"bubble","title":"labelPos=top","dataLabels":"value","labelPos":"top",
       "legend":"none","data":D})
ch(BR,{"chartType":"bubble","title":"dataLabels=none","dataLabels":"none","legend":"none","data":D})

new_slide("Axes — min/max, gridlines, ticks")
ch(TL,{"chartType":"bubble","title":"min/max + titles","axismin":"0","axismax":"30",
       "majorunit":"10","axistitle":"Y","cattitle":"X","axisfont":"10:333333:Calibri",
       "axisline":"666666:1","legend":"none","data":D})
ch(TR,{"chartType":"bubble","title":"gridlines + minorGridlines",
       "gridlines":"E0E0E0:0.3","minorGridlines":"F0F0F0:0.25","legend":"none","data":D})
ch(BL,{"chartType":"bubble","title":"labelrotation=-30","labelrotation":"-30","legend":"none","data":D})
ch(BR,{"chartType":"bubble","title":"dispunits=hundreds","dispunits":"hundreds","legend":"none",
       "data":"A:500,1200,800,1800,2200,900"})

new_slide("Series styling — colors, gradient, transparency, outline, shadow")
ch(TL,{"chartType":"bubble","title":"colors + seriesoutline","colors":"4472C4,ED7D31",
       "seriesoutline":"000000:0.5","legend":"bottom","data":D2})
ch(TR,{"chartType":"bubble","title":"gradient + seriesshadow",
       "gradient":"FF6600-FFCC00","seriesshadow":"000000-5-45-3-50",
       "legend":"none","data":D})
ch(BL,{"chartType":"bubble","title":"transparency=30","transparency":"30",
       "legend":"bottom","data":D2})
ch(BR,{"chartType":"bubble","title":"per-series gradients",
       "gradients":"FF0000-0000FF;00FF00-FFFF00","legend":"bottom","data":D2})

new_slide("Presets & per-series Set")
for box,p in zip([TL,TR,BL],["minimal","dark","corporate"]):
    ch(box,{"chartType":"bubble","preset":p,"title":f"preset={p}","legend":"bottom","data":D2})
ch(BR,{"chartType":"bubble","title":"chart-series Set name+color","legend":"bottom","data":D2})
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[1]",
    *P({"name":"Renamed A","color":"C00000"}))
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[2]",
    *P({"name":"Renamed B","color":"2E75B6"}))

print(f"Done: {FILE}  ({slide} slides)")
