#!/usr/bin/env python3
"""
Stock Charts Showcase — High-Low-Close and OHLC variants.

Generates: charts-stock.pptx

  Slide 1  Basic stock         3-series HLC + 4-series OHLC
  Slide 2  Hi-low / up-down    hilowlines, updownbars
  Slide 3  Title & legend
  Slide 4  Data labels
  Slide 5  Axes                min/max, gridlines, axisnumfmt (currency)
  Slide 6  Series styling      colors, transparency, outline, shadow
  Slide 7  Backgrounds         chartareafill, plotFill, chartborder
  Slide 8  Presets & per-ser   preset bundles + chart-series Set

Usage:
  python3 charts-stock.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-stock.pptx")
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
HLC="High:130,135,140,138,145;Low:118,122,128,125,132;Close:125,130,135,132,140"
OHLC="Open:120,128,130,135,138;High:130,135,140,138,145;Low:118,122,128,125,132;Close:125,130,135,132,140"

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

new_slide("Basic stock — High-Low-Close vs Open-High-Low-Close")
ch(TL,{"chartType":"stock","title":"HLC","legend":"bottom","categories":CATS,"data":HLC})
ch(TR,{"chartType":"stock","title":"OHLC","legend":"bottom","categories":CATS,"data":OHLC})
ch(BL,{"chartType":"stock","title":"HLC + dataTable=true","dataTable":"true",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(BR,{"chartType":"stock","title":"OHLC + dataTable=true","dataTable":"true",
       "legend":"bottom","categories":CATS,"data":OHLC})

new_slide("hilowlines & updownbars")
ch(TL,{"chartType":"stock","title":"hilowlines=true","hilowlines":"true",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(TR,{"chartType":"stock","title":"hilowlines=808080:0.5","hilowlines":"808080:0.5",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(BL,{"chartType":"stock","title":"updownbars=true (OHLC)","updownbars":"true",
       "legend":"bottom","categories":CATS,"data":OHLC})
ch(BR,{"chartType":"stock","title":"updownbars=150:00AA00:FF0000",
       "updownbars":"150:00AA00:FF0000","legend":"bottom","categories":CATS,"data":OHLC})

new_slide("Title & legend")
ch(TL,{"chartType":"stock","title":"Styled title","title.font":"Georgia","title.size":"20",
       "title.color":"4472C4","title.bold":"true","legend":"bottom","categories":CATS,"data":HLC})
ch(TR,{"chartType":"stock","title":"legend=top + legendFont","legend":"top",
       "legendFont":"10:333333:Calibri","categories":CATS,"data":HLC})
ch(BL,{"chartType":"stock","title":"legend.overlay=true","legend":"topRight",
       "legend.overlay":"true","categories":CATS,"data":HLC})
ch(BR,{"chartType":"stock","autotitledeleted":"true","legend":"none","categories":CATS,"data":HLC})

new_slide("Data labels — flags + labelfont")
ch(TL,{"chartType":"stock","title":"dataLabels=value","dataLabels":"value",
       "labelfont":"9:333333:Calibri","legend":"bottom","categories":CATS,"data":HLC})
ch(TR,{"chartType":"stock","title":"value,series","dataLabels":"value,series",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(BL,{"chartType":"stock","title":"value,category","dataLabels":"value,category",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(BR,{"chartType":"stock","title":"dataLabels=none","dataLabels":"none",
       "legend":"bottom","categories":CATS,"data":HLC})

new_slide("Axes — min/max, gridlines, currency format")
ch(TL,{"chartType":"stock","title":"min/max + titles","axismin":"100","axismax":"160",
       "majorunit":"10","axistitle":"Price (USD)","cattitle":"Day",
       "axisfont":"10:333333:Calibri","axisnumfmt":"$#,##0.00",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(TR,{"chartType":"stock","title":"gridlines + minorGridlines",
       "gridlines":"E0E0E0:0.3","minorGridlines":"F0F0F0:0.25",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(BL,{"chartType":"stock","title":"labelrotation=-30","labelrotation":"-30",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(BR,{"chartType":"stock","title":"dispunits=hundreds","dispunits":"hundreds",
       "legend":"bottom","categories":CATS,
       "data":"High:13000,13500,14000,13800,14500;Low:11800,12200,12800,12500,13200;Close:12500,13000,13500,13200,14000"})

new_slide("Series styling — colors, transparency, outline, shadow")
ch(TL,{"chartType":"stock","title":"colors","colors":"4472C4,ED7D31,70AD47",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(TR,{"chartType":"stock","title":"seriesoutline","seriesoutline":"000000:1",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(BL,{"chartType":"stock","title":"transparency=30","transparency":"30",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(BR,{"chartType":"stock","title":"seriesshadow","seriesshadow":"000000-5-45-3-50",
       "legend":"bottom","categories":CATS,"data":HLC})

new_slide("Backgrounds — chartareafill, plotFill, chartborder, roundedcorners")
ch(TL,{"chartType":"stock","title":"chartareafill + plotFill + borders",
       "chartareafill":"FFF8E7","plotFill":"FAFAFA","chartborder":"000000:1",
       "plotborder":"CCCCCC:0.5","legend":"bottom","categories":CATS,"data":HLC})
ch(TR,{"chartType":"stock","title":"roundedcorners=true","roundedcorners":"true",
       "chartborder":"4472C4:2","legend":"bottom","categories":CATS,"data":HLC})
ch(BL,{"chartType":"stock","title":"plotFill=none","plotFill":"none","gridlines":"none",
       "legend":"bottom","categories":CATS,"data":HLC})
ch(BR,{"chartType":"stock","title":"chartareafill=none","chartareafill":"none",
       "legend":"bottom","categories":CATS,"data":HLC})

new_slide("Presets & per-series Set")
for box,p in zip([TL,TR,BL],["minimal","dark","corporate"]):
    ch(box,{"chartType":"stock","preset":p,"title":f"preset={p}","legend":"bottom",
            "categories":CATS,"data":HLC})
ch(BR,{"chartType":"stock","title":"chart-series Set name+color","legend":"bottom",
       "categories":CATS,"data":HLC})
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[1]",*P({"name":"H","color":"00AA00"}))
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[2]",*P({"name":"L","color":"C00000"}))
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[3]",*P({"name":"C","color":"4472C4"}))

print(f"Done: {FILE}  ({slide} slides)")
