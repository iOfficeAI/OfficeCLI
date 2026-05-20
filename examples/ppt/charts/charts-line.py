#!/usr/bin/env python3
"""
Line Charts Showcase — line, stackedLine, percentStackedLine, line3d.

Generates: charts-line.pptx

  Slide 1  Variants           line / stackedLine / percentStackedLine / line3d
  Slide 2  Markers            marker symbol/size/color, markersize, showMarker
  Slide 3  Smoothing & dash   smooth, linedash, linewidth
  Slide 4  Title & legend     title.* + legend positions + legendFont
  Slide 5  Data labels        flags, labelPos, labelfont
  Slide 6  Axes               min/max, titles, fonts, gridlines, ticks, labelrotation, log
  Slide 7  Overlays           droplines, hilowlines, updownbars, trendline, errbars, referenceline
  Slide 8  Per-series Set     lineWidth/lineDash/marker/markerSize/color/smooth + presets

Usage:
  python3 charts-line.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-line.pptx")
def cli(*a):
    r = subprocess.run(["officecli", *a], capture_output=True, text=True)
    if r.returncode:
        m = (r.stderr or r.stdout or "").strip().splitlines()
        head = m[0][:160] if m else ""
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
CATS="Mon,Tue,Wed,Thu,Fri"; D2="A:50,60,70,65,80;B:40,45,55,60,75"

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

new_slide("Line variants — line / stackedLine / percentStackedLine / line3d")
ch(TL,{"chartType":"line","title":"line","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"stackedLine","title":"stackedLine","legend":"bottom","categories":CATS,"data":D2})
ch(BL,{"chartType":"percentStackedLine","title":"percentStackedLine","legend":"bottom","categories":CATS,"data":D2})
ch(BR,{"chartType":"line3d","title":"line3d","legend":"bottom","categories":CATS,"data":D2})

new_slide("Markers — symbol, size, color, showMarker")
ch(TL,{"chartType":"line","title":"marker=circle:8:FF0000","marker":"circle:8:FF0000",
       "linewidth":"2","legend":"none","categories":CATS,"data":"A:50,60,70,65,80"})
ch(TR,{"chartType":"line","title":"marker=square:6","marker":"square:6","linewidth":"2",
       "legend":"none","categories":CATS,"data":"A:50,60,70,65,80"})
ch(BL,{"chartType":"line","title":"marker=diamond:10:0070C0","marker":"diamond:10:0070C0",
       "linewidth":"2","legend":"none","categories":CATS,"data":"A:50,60,70,65,80"})
ch(BR,{"chartType":"line","title":"showMarker=true (default markers)","showMarker":"true",
       "legend":"bottom","categories":CATS,"data":D2})

new_slide("Smoothing & dash — smooth, linedash, linewidth")
ch(TL,{"chartType":"line","title":"smooth=true","smooth":"true","linewidth":"2.5",
       "legend":"none","categories":CATS,"data":"A:50,60,70,65,80"})
ch(TR,{"chartType":"line","title":"linedash=dash","linedash":"dash","linewidth":"2",
       "legend":"none","categories":CATS,"data":"A:50,60,70,65,80"})
ch(BL,{"chartType":"line","title":"linedash=dot","linedash":"dot","linewidth":"2",
       "legend":"none","categories":CATS,"data":"A:50,60,70,65,80"})
ch(BR,{"chartType":"line","title":"linedash=dashDot","linedash":"dashDot","linewidth":"2",
       "legend":"none","categories":CATS,"data":"A:50,60,70,65,80"})

new_slide("Title & legend")
ch(TL,{"chartType":"line","title":"Styled title","title.font":"Georgia","title.size":"20",
       "title.color":"4472C4","title.bold":"true","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"line","title":"legend=top + legendFont","legend":"top",
       "legendFont":"10:333333:Calibri","categories":CATS,"data":D2})
ch(BL,{"chartType":"line","title":"legend.overlay=true","legend":"topRight",
       "legend.overlay":"true","categories":CATS,"data":D2})
ch(BR,{"chartType":"line","autotitledeleted":"true","legend":"none","categories":CATS,"data":D2})

new_slide("Data labels — flags, labelPos, labelfont")
ch(TL,{"chartType":"line","title":"dataLabels=value @ top","dataLabels":"value","labelPos":"top",
       "labelfont":"10:333333:Calibri","legend":"none","categories":CATS,"data":"A:50,60,70,65,80"})
ch(TR,{"chartType":"line","title":"value,category","dataLabels":"value,category","labelPos":"top",
       "legend":"none","categories":CATS,"data":"A:50,60,70,65,80"})
ch(BL,{"chartType":"line","title":"dataLabels=none","dataLabels":"none","legend":"none",
       "categories":CATS,"data":"A:50,60,70,65,80"})
ch(BR,{"chartType":"line","title":"labelfont styled","dataLabels":"value","labelPos":"top",
       "labelfont":"12:C00000:Georgia","legend":"none","categories":CATS,"data":"A:50,60,70,65,80"})

new_slide("Axes — min/max, gridlines, ticks, labelrotation, log")
ch(TL,{"chartType":"line","title":"min/max + titles","legend":"none",
       "axismin":"0","axismax":"100","majorunit":"25","axistitle":"Visits","cattitle":"Day",
       "axisfont":"10:333333:Calibri","axisline":"666666:1","axisnumfmt":"#,##0",
       "categories":CATS,"data":"A:50,60,70,65,80"})
ch(TR,{"chartType":"line","title":"gridlines + ticks","legend":"none",
       "gridlines":"E0E0E0:0.3","minorGridlines":"F0F0F0:0.25",
       "majorTickMark":"out","minorTickMark":"in","tickLabelPos":"nextTo",
       "categories":CATS,"data":"A:50,60,70,65,80"})
ch(TR if False else BL,{"chartType":"line","title":"labelrotation=-30","legend":"none",
       "labelrotation":"-30","categories":"January,February,March,April,May,June",
       "data":"A:60,90,140,180,160,210"})
ch(BR,{"chartType":"line","title":"logbase=10","legend":"none","logbase":"10",
       "axismin":"1","axismax":"10000","categories":CATS,"data":"Growth:5,50,500,5000,3000"})

new_slide("Overlays — droplines, hilowlines, updownbars, trendline, errbars, referenceline")
ch(TL,{"chartType":"line","title":"droplines + hilowlines","droplines":"808080:0.5","hilowlines":"true",
       "legend":"bottom","categories":CATS,"data":"High:130,135,140,138,145;Low:118,122,128,125,132"})
ch(TR,{"chartType":"line","title":"updownbars=150:00AA00:FF0000",
       "updownbars":"150:00AA00:FF0000","legend":"bottom","categories":CATS,
       "data":"Open:120,128,130,135,138;Close:128,125,135,138,142"})
ch(BL,{"chartType":"line","title":"trendline=linear + errbars=stdDev:1",
       "trendline":"linear","errbars":"stdDev:1","legend":"none",
       "categories":CATS,"data":"A:50,60,70,65,80"})
ch(BR,{"chartType":"line","title":"referenceline=70:FF0000:Target",
       "referenceline":"70:FF0000:Target","legend":"none",
       "categories":CATS,"data":"A:50,60,70,65,80"})

new_slide("Per-series Set + presets — chart-series lineWidth/lineDash/marker/markerSize/color/smooth")
ch(TL,{"chartType":"line","preset":"minimal","title":"preset=minimal",
       "legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"line","preset":"dark","title":"preset=dark",
       "legend":"bottom","categories":CATS,"data":D2})
ch(BL,{"chartType":"line","preset":"corporate","title":"preset=corporate",
       "legend":"bottom","categories":CATS,"data":D2})
ch(BR,{"chartType":"line","title":"chart-series Set per line","showMarker":"true",
       "legend":"bottom","categories":CATS,"data":D2})
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[1]",
    *P({"name":"Alpha","color":"C00000","lineWidth":"2.5","lineDash":"solid",
        "marker":"circle","markerSize":"9","smooth":"true"}))
cli("set",FILE,f"/slide[{slide}]/chart[4]/series[2]",
    *P({"name":"Beta","color":"2E75B6","lineWidth":"1.5","lineDash":"dash",
        "marker":"diamond","markerSize":"8"}))

print(f"Done: {FILE}  ({slide} slides)")
