#!/usr/bin/env python3
"""
Bar Charts Showcase — bar, stackedBar, percentStackedBar, bar3d (cylinder/cone/pyramid).

Generates: charts-bar.pptx

  Slide 1  Variants           bar / stackedBar / percentStackedBar / bar3d
  Slide 2  3D bar shapes      shape=box/cylinder/cone/pyramid (bar3d only)
  Slide 3  Title & legend     title.* + legend positions + legendFont
  Slide 4  Data labels        flags + labelPos + labelfont
  Slide 5  Axes               min/max/title/font/line/numfmt/gridlines/labelrotation
  Slide 6  Series styling     colors, gradient, transparency, outline, shadow, invertifneg, serlines
  Slide 7  Overlays           referenceline, errbars, gapwidth, overlap, dataTable
  Slide 8  Presets & per-ser  preset bundles + seriesN.* + chart-series Set

Usage:
  python3 charts-bar.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-bar.pptx")
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
slide = 0
def new_slide(t):
    global slide; slide += 1
    cli("add", FILE, "/", "--type", "slide")
    cli("add", FILE, f"/slide[{slide}]", "--type", "shape",
        *P({"text": t, "size": 24, "bold": "true",
            "autoFit":"normal","x":"0.5in","y":"0.3in","width":"12.3in","height":"0.6in"}))
def ch(box, p): cli("add", FILE, f"/slide[{slide}]", "--type", "chart", *P({**box, **p}))
TL={"x":"0.3in","y":"1.05in","width":"6.1in","height":"3in"}
TR={"x":"6.95in","y":"1.05in","width":"6.1in","height":"3in"}
BL={"x":"0.3in","y":"4.25in","width":"6.1in","height":"3in"}
BR={"x":"6.95in","y":"4.25in","width":"6.1in","height":"3in"}
CATS="Q1,Q2,Q3,Q4"; D2="East:120,135,148,162;West:95,108,115,128"
D3="East:120,135,148,162;South:95,108,115,128;West:80,90,98,110"

if os.path.exists(FILE): os.remove(FILE)
cli("create", FILE); cli("open", FILE)
atexit.register(lambda: (cli("close", FILE), cli("validate", FILE)))

new_slide("Bar variants — bar / stackedBar / percentStackedBar / bar3d")
ch(TL,{"chartType":"bar","title":"bar","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"stackedBar","title":"stackedBar","legend":"bottom","categories":CATS,"data":D3})
ch(BL,{"chartType":"percentStackedBar","title":"percentStackedBar","legend":"bottom","categories":CATS,"data":D3})
ch(BR,{"chartType":"bar3d","title":"bar3d","legend":"bottom","categories":CATS,"data":D2,"view3d":"15,20,30"})

new_slide("3D bar shapes — shape=box / cylinder / cone / pyramid")
ch(TL,{"chartType":"bar3d","shape":"box","title":"shape=box","legend":"none","categories":CATS,"data":D2})
ch(TR,{"chartType":"bar3d","shape":"cylinder","title":"shape=cylinder","legend":"none","categories":CATS,"data":D2})
ch(BL,{"chartType":"bar3d","shape":"cone","title":"shape=cone","legend":"none","categories":CATS,"data":D2})
ch(BR,{"chartType":"bar3d","shape":"pyramid","title":"shape=pyramid","legend":"none","categories":CATS,"data":D2})

new_slide("Title & legend")
ch(TL,{"chartType":"bar","title":"Styled title","title.font":"Georgia","title.size":"20",
       "title.color":"4472C4","title.bold":"true","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"bar","title":"legend=top + legendFont","legend":"top",
       "legendFont":"10:333333:Calibri","categories":CATS,"data":D2})
ch(BL,{"chartType":"bar","title":"legend.overlay=true","legend":"topRight",
       "legend.overlay":"true","categories":CATS,"data":D2})
ch(BR,{"chartType":"bar","autotitledeleted":"true","legend":"none","categories":CATS,"data":D2})

new_slide("Data labels — flags, labelPos, labelfont")
ch(TL,{"chartType":"bar","title":"value @ outsideEnd","dataLabels":"value",
       "labelPos":"outsideEnd","labelfont":"10:333333:Calibri","legend":"none",
       "categories":CATS,"data":"A:60,90,140,180"})
ch(TR,{"chartType":"bar","title":"value,category @ insideEnd","dataLabels":"value,category",
       "labelPos":"insideEnd","labelfont":"9:FFFFFF:Calibri","legend":"none",
       "categories":CATS,"data":"A:60,90,140,180"})
ch(BL,{"chartType":"stackedBar","title":"stacked + center labels","dataLabels":"value",
       "labelPos":"center","labelfont":"9:FFFFFF:Calibri","legend":"bottom",
       "categories":CATS,"data":D3})
ch(BR,{"chartType":"bar","title":"dataLabels=none","dataLabels":"none","legend":"none",
       "categories":CATS,"data":"A:60,90,140,180"})

new_slide("Axes — min/max, titles, fonts, gridlines, ticks, labelrotation")
ch(TL,{"chartType":"bar","title":"min/max + titles + numfmt","legend":"none",
       "axismin":"0","axismax":"200","majorunit":"50","minorunit":"10",
       "axistitle":"Revenue","cattitle":"Quarter","axisfont":"10:333333:Calibri",
       "axisline":"666666:1","axisnumfmt":"#,##0","categories":CATS,"data":"Rev:60,90,140,180"})
ch(TR,{"chartType":"bar","title":"gridlines + ticks","legend":"none",
       "gridlines":"E0E0E0:0.3","minorGridlines":"F0F0F0:0.25",
       "majorTickMark":"out","minorTickMark":"in","tickLabelPos":"nextTo",
       "categories":CATS,"data":"A:60,90,140,180"})
ch(BL,{"chartType":"bar","title":"labelrotation=-30","legend":"none","labelrotation":"-30",
       "categories":"January,February,March,April","data":"A:60,90,140,180"})
ch(BR,{"chartType":"bar","title":"dispunits=thousands","legend":"none","dispunits":"thousands",
       "categories":CATS,"data":"Rev:120000,135000,148000,162000"})
cli("set", FILE, f"/slide[{slide}]/chart[1]/axis[@role=value]",
    *P({"title":"Revenue","format":"$#,##0","majorGridlines":"true","max":"200","min":"0"}))

new_slide("Series styling — colors, gradient(s), transparency, outline, shadow, invertifneg, serlines")
ch(TL,{"chartType":"bar","title":"colors + seriesoutline","legend":"bottom",
       "colors":"4472C4,ED7D31,A5A5A5","seriesoutline":"000000:0.5","categories":CATS,"data":D3})
ch(TR,{"chartType":"bar","title":"gradient + seriesshadow","legend":"bottom",
       "gradient":"FF6600-FFCC00:90","seriesshadow":"000000-5-45-3-50",
       "categories":CATS,"data":"A:60,90,140,180"})
ch(BL,{"chartType":"bar","title":"transparency=30 + gradients","legend":"bottom",
       "gradients":"FF0000-0000FF;00FF00-FFFF00","transparency":"30",
       "categories":CATS,"data":"A:60,90,140,180;B:40,70,100,130"})
ch(BR,{"chartType":"stackedBar","title":"stacked + serlines=true","serlines":"true",
       "legend":"bottom","categories":CATS,"data":D2})

new_slide("Overlays — referenceline, errbars, gapwidth, overlap, dataTable")
ch(TL,{"chartType":"bar","title":"referenceline=100","legend":"none",
       "referenceline":"100:FF0000:Target","categories":CATS,"data":"A:60,90,140,180"})
ch(TR,{"chartType":"bar","title":"errbars=fixedVal:10","legend":"none",
       "errbars":"fixedVal:10","categories":CATS,"data":"A:60,90,140,180"})
ch(BL,{"chartType":"bar","title":"gapwidth=50 + overlap=20","legend":"bottom",
       "gapwidth":"50","overlap":"20","categories":CATS,
       "data":"A:60,90,140,180;B:50,75,110,150"})
ch(BR,{"chartType":"bar","title":"dataTable=true","legend":"bottom",
       "dataTable":"true","categories":CATS,"data":"A:60,90,140,180"})

new_slide("Presets & per-series control")
ch(TL,{"chartType":"bar","preset":"minimal","title":"preset=minimal","legend":"bottom",
       "categories":CATS,"data":"A:60,90,140,180;B:50,75,110,150"})
ch(TR,{"chartType":"bar","preset":"dark","title":"preset=dark","legend":"bottom",
       "categories":CATS,"data":"A:60,90,140,180;B:50,75,110,150"})
ch(BL,{"chartType":"bar","preset":"corporate","title":"preset=corporate","legend":"bottom",
       "categories":CATS,"data":"A:60,90,140,180;B:50,75,110,150"})
ch(BR,{"chartType":"bar","title":"seriesN.* Add + chart-series Set","legend":"bottom",
       "categories":CATS,
       "series1.name":"Product A","series1.values":"60,90,140,180","series1.color":"4472C4",
       "series2.name":"Product B","series2.values":"50,75,110,150","series2.color":"ED7D31"})
cli("set", FILE, f"/slide[{slide}]/chart[4]/series[1]", *P({"name":"Renamed","color":"C00000"}))

print(f"Done: {FILE}  ({slide} slides)")
