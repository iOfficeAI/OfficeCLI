#!/usr/bin/env python3
"""
Waterfall Charts Showcase — increaseColor / decreaseColor / totalColor.

Generates: charts-waterfall.pptx

  Slide 1  Basic                  default colors, single dataset
  Slide 2  Color schemes          increaseColor / decreaseColor / totalColor combinations
  Slide 3  Title & legend
  Slide 4  Data labels
  Slide 5  Axes                   min/max, gridlines, axisnumfmt (currency)
  Slide 6  Backgrounds            chartareafill, plotFill, chartborder, roundedcorners
  Slide 7  Larger story           a real cashflow waterfall with labels
  Slide 8  Presets

Usage:
  python3 charts-waterfall.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-waterfall.pptx")
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
HERO={"x":"1in","y":"1.05in","width":"11.3in","height":"6.2in"}
CATS="Start,Q1,Q2,Q3,Q4,End"
D="Cashflow:100,30,-15,40,-10,145"
CATS_LONG="Open,Revenue,COGS,Opex,R&D,Tax,Net"
D_LONG="P&L:100,80,-30,-25,-15,-10,100"

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

new_slide("Basic waterfall — default colors")
ch(TL,{"chartType":"waterfall","title":"Default colors","legend":"none",
       "categories":CATS,"data":D})
ch(TR,{"chartType":"waterfall","title":"Default + dataTable","dataTable":"true",
       "legend":"none","categories":CATS,"data":D})
ch(BL,{"chartType":"waterfall","title":"With legend","legend":"bottom",
       "categories":CATS,"data":D})
ch(BR,{"chartType":"waterfall","title":"7-step P&L","legend":"none",
       "categories":CATS_LONG,"data":D_LONG})

new_slide("Color schemes — increaseColor / decreaseColor / totalColor")
ch(TL,{"chartType":"waterfall","title":"green/red/blue (default-ish)",
       "increaseColor":"00AA00","decreaseColor":"FF0000","totalColor":"4472C4",
       "legend":"none","categories":CATS,"data":D})
ch(TR,{"chartType":"waterfall","title":"corporate (teal/orange/navy)",
       "increaseColor":"008080","decreaseColor":"D86600","totalColor":"1F3864",
       "legend":"none","categories":CATS,"data":D})
ch(BL,{"chartType":"waterfall","title":"monochrome",
       "increaseColor":"606060","decreaseColor":"A0A0A0","totalColor":"303030",
       "legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"waterfall","title":"vivid",
       "increaseColor":"00C853","decreaseColor":"D50000","totalColor":"2962FF",
       "legend":"none","categories":CATS,"data":D})

new_slide("Title & legend")
ch(TL,{"chartType":"waterfall","title":"Styled title","title.font":"Georgia","title.size":"20",
       "title.color":"4472C4","title.bold":"true","legend":"bottom","categories":CATS,"data":D})
ch(TR,{"chartType":"waterfall","title":"legend=top + legendFont","legend":"top",
       "legendFont":"10:333333:Calibri","categories":CATS,"data":D})
ch(BL,{"chartType":"waterfall","title":"legend.overlay=true","legend":"topRight",
       "legend.overlay":"true","categories":CATS,"data":D})
ch(BR,{"chartType":"waterfall","autotitledeleted":"true","legend":"none",
       "categories":CATS,"data":D})

new_slide("Data labels — flags + labelfont")
ch(TL,{"chartType":"waterfall","title":"value","dataLabels":"value",
       "labelfont":"10:333333:Calibri","legend":"none","categories":CATS,"data":D})
ch(TR,{"chartType":"waterfall","title":"value,category","dataLabels":"value,category",
       "legend":"none","categories":CATS,"data":D})
ch(BL,{"chartType":"waterfall","title":"value @ outsideEnd","dataLabels":"value",
       "labelPos":"outsideEnd","legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"waterfall","title":"dataLabels=none","dataLabels":"none",
       "legend":"none","categories":CATS,"data":D})

new_slide("Axes — min/max, titles, gridlines, axisnumfmt")
ch(TL,{"chartType":"waterfall","title":"min/max + titles","axismin":"0","axismax":"200",
       "majorunit":"50","axistitle":"USD","cattitle":"Phase",
       "axisfont":"10:333333:Calibri","axisnumfmt":"$#,##0",
       "legend":"none","categories":CATS,"data":D})
ch(TR,{"chartType":"waterfall","title":"gridlines + minorGridlines",
       "gridlines":"E0E0E0:0.3","minorGridlines":"F0F0F0:0.25",
       "legend":"none","categories":CATS,"data":D})
ch(BL,{"chartType":"waterfall","title":"labelrotation=-30","labelrotation":"-30",
       "legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"waterfall","title":"dispunits=thousands","dispunits":"thousands",
       "legend":"none","categories":CATS,
       "data":"USD:100000,30000,-15000,40000,-10000,145000"})

new_slide("Backgrounds — chartareafill, plotFill, chartborder, roundedcorners")
ch(TL,{"chartType":"waterfall","title":"chartareafill + chartborder",
       "chartareafill":"FFF8E7","chartborder":"000000:1","plotFill":"FAFAFA",
       "plotborder":"CCCCCC:0.5","legend":"none","categories":CATS,"data":D})
ch(TR,{"chartType":"waterfall","title":"roundedcorners=true","roundedcorners":"true",
       "chartborder":"4472C4:2","legend":"none","categories":CATS,"data":D})
ch(BL,{"chartType":"waterfall","title":"plotFill=none","plotFill":"none",
       "gridlines":"none","legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"waterfall","title":"chartareafill=none","chartareafill":"none",
       "legend":"none","categories":CATS,"data":D})

new_slide("Hero cashflow waterfall — full slide with labels")
ch(HERO,{"chartType":"waterfall","title":"FY24 P&L Walk",
         "title.font":"Helvetica","title.size":"22","title.bold":"true","title.color":"1F3864",
         "increaseColor":"00C853","decreaseColor":"D50000","totalColor":"2962FF",
         "dataLabels":"value,category","labelPos":"outsideEnd",
         "labelfont":"11:333333:Helvetica","axistitle":"USD","cattitle":"",
         "axisnumfmt":"$#,##0","gridlines":"E0E0E0:0.3",
         "legend":"none","categories":CATS_LONG,"data":D_LONG})

new_slide("Presets")
for box,p in zip([TL,TR,BL,BR],["minimal","dark","corporate","colorful"]):
    ch(box,{"chartType":"waterfall","preset":p,"title":f"preset={p}",
            "legend":"none","categories":CATS,"data":D})

print(f"Done: {FILE}  ({slide} slides)")
