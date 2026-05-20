#!/usr/bin/env python3
"""
Advanced Charts Showcase — properties not covered by the per-type decks.

Generates: charts-advanced.pptx

Coverage of the long tail of chart properties (cross-handler / niche / axis-level):

  Slide 1  RTL & anchor          direction=rtl, anchor named-token, anchor cm-form
  Slide 2  Axis-level shortcuts  axisvisible / valaxisvisible / catAxisVisible,
                                 axisorientation, axisposition,
                                 cataxisline / valaxisline
  Slide 3  Crossings             crossBetween (between/midCat), crosses (autoZero/max/min), crossesAt
  Slide 4  Categories axis       labeloffset, ticklabelskip
  Slide 5  Marker size & fills   markersize (standalone), areafill, chartFill, plotvisonly
  Slide 6  Built-in style + blanks  style=1..48, dispBlanksAs (gap / zero / span)
  Slide 7  chart-axis Set        dispUnits, logBase, minorUnit, visible, labelRotation (per-axis)
  Slide 8  chart-series mutation values=, categories= (per-series range), + get-readback round-trip

Usage:
  python3 charts-advanced.py
"""
import subprocess, os, sys, atexit, json

FILE = os.path.join(os.path.dirname(__file__), "charts-advanced.pptx")
def cli(*a, capture=False):
    r = subprocess.run(["officecli", *a], capture_output=True, text=True)
    if r.returncode:
        m=(r.stderr or r.stdout or "").strip().splitlines()
        head=m[0][:160] if m else ""
        if "UNSUPPORTED" in (r.stderr or ""):
            # Forward-compat: skip unsupported props but surface so silent gaps are visible.
            print(f"  ⚠ {' '.join(a[:3])} → {head}", file=sys.stderr); return
        if m: print(f"  ! {' '.join(a[:3])} → {head}", file=sys.stderr)
        sys.exit(r.returncode)
    if capture: return r.stdout
def P(d): return [x for k,v in d.items() for x in ("--prop", f"{k}={v}")]
slide=0
def new_slide(t):
    global slide; slide+=1
    cli("add",FILE,"/","--type","slide")
    cli("add",FILE,f"/slide[{slide}]","--type","shape",
        *P({"text":t,"size":24,"bold":"true","autoFit":"normal","x":"0.5in","y":"0.3in","width":"12.3in","height":"0.6in"}))
def ch(box,p): cli("add",FILE,f"/slide[{slide}]","--type","chart",*P({**box,**p}))
def note(x,y,text):
    cli("add",FILE,f"/slide[{slide}]","--type","shape",
        *P({"text":text,"size":10,"italic":"true","color":"666666",
            "x":x,"y":y,"width":"6in","height":"0.4in"}))
TL={"x":"0.3in","y":"1.05in","width":"6.1in","height":"3in"}
TR={"x":"6.95in","y":"1.05in","width":"6.1in","height":"3in"}
BL={"x":"0.3in","y":"4.25in","width":"6.1in","height":"3in"}
BR={"x":"6.95in","y":"4.25in","width":"6.1in","height":"3in"}
CATS="Q1,Q2,Q3,Q4"
D="A:60,90,140,180"
D2="A:60,90,140,180;B:50,75,110,150"

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

# ---------------------------------------------------------------------------
# Slide 1 — RTL + anchor variants
# ---------------------------------------------------------------------------
new_slide("RTL + anchor — direction=rtl, named-token anchor, cm-form anchor")
ch(TL,{"chartType":"column","title":"default (LTR)","legend":"bottom",
       "categories":CATS,"data":D2})
# RTL must be Set after Add (direction is set-only)
ch(TR,{"chartType":"column","title":"direction=rtl (Set after Add)","legend":"bottom",
       "categories":"Q1,Q2,Q3,Q4","data":D2})
cli("set",FILE,f"/slide[{slide}]/chart[2]",*P({"direction":"rtl"}))
# Anchor cm-form: x,y,w,h
ch({"anchor":"0.3cm,11cm,15.5cm,7cm"},{"chartType":"column",
    "title":"anchor=0.3cm,11cm,15.5cm,7cm","legend":"bottom",
    "categories":CATS,"data":D})

# ---------------------------------------------------------------------------
# Slide 2 — axis-level shortcuts
# ---------------------------------------------------------------------------
new_slide("Axis shortcuts — axisvisible / valaxisvisible / catAxisVisible, orientation, position, lines")
ch(TL,{"chartType":"column","title":"axisvisible=false (both axes hidden)",
       "legend":"none","axisvisible":"false","categories":CATS,"data":D})
ch(TR,{"chartType":"column","title":"valaxisvisible=false (Y hidden, X shown)",
       "legend":"none","valaxisvisible":"false","categories":CATS,"data":D})
ch(BL,{"chartType":"column","title":"catAxisVisible=false (X hidden)",
       "legend":"none","catAxisVisible":"false","categories":CATS,"data":D})
ch(BR,{"chartType":"column","title":"axisorientation=true (reversed) + axisposition=top",
       "legend":"none","axisorientation":"true","axisposition":"top",
       "cataxisline":"333333:1","valaxisline":"333333:1",
       "categories":CATS,"data":D})

# ---------------------------------------------------------------------------
# Slide 3 — Crossings
# ---------------------------------------------------------------------------
new_slide("Crossings — crossBetween / crosses / crossesAt")
ch(TL,{"chartType":"column","title":"crossBetween=between (default)",
       "legend":"none","crossBetween":"between","categories":CATS,"data":D})
ch(TR,{"chartType":"column","title":"crossBetween=midCat","legend":"none",
       "crossBetween":"midCat","categories":CATS,"data":D})
ch(BL,{"chartType":"column","title":"crosses=max (Y crosses at top)","legend":"none",
       "crosses":"max","categories":CATS,"data":D})
ch(BR,{"chartType":"column","title":"crossesAt=100 + crosses=autoZero",
       "legend":"none","crosses":"autoZero","crossesAt":"100",
       "categories":CATS,"data":"A:60,-30,140,180"})

# ---------------------------------------------------------------------------
# Slide 4 — Category axis layout
# ---------------------------------------------------------------------------
new_slide("Category axis — labeloffset, ticklabelskip")
ch(TL,{"chartType":"column","title":"labeloffset=100 (default)",
       "labeloffset":"100","legend":"none",
       "categories":"January,February,March,April,May,June",
       "data":"A:60,90,140,180,160,210"})
ch(TR,{"chartType":"column","title":"labeloffset=300 (push labels down)",
       "labeloffset":"300","legend":"none",
       "categories":"January,February,March,April,May,June",
       "data":"A:60,90,140,180,160,210"})
ch(BL,{"chartType":"column","title":"ticklabelskip=2 (every other label)",
       "ticklabelskip":"2","legend":"none",
       "categories":"Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec",
       "data":"A:60,90,140,180,160,210,200,190,170,150,130,170"})
ch(BR,{"chartType":"column","title":"ticklabelskip=3","ticklabelskip":"3","legend":"none",
       "categories":"Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec",
       "data":"A:60,90,140,180,160,210,200,190,170,150,130,170"})

# ---------------------------------------------------------------------------
# Slide 5 — Marker size, area/chart fills, plotvisonly
# ---------------------------------------------------------------------------
new_slide("Marker size & fills — markersize (standalone), areafill, chartFill, plotvisonly")
ch(TL,{"chartType":"line","title":"markersize=12 (standalone key)",
       "showMarker":"true","markersize":"12","legend":"none",
       "categories":CATS,"data":D})
ch(TR,{"chartType":"column","title":"areafill (applies to every series shape)",
       "areafill":"4472C4-A5C8FF:90","legend":"none","categories":CATS,"data":D2})
ch(BL,{"chartType":"column","title":"chartFill=#FFF8E7 (chart-level fill)",
       "chartFill":"#FFF8E7","legend":"none","categories":CATS,"data":D})
ch(BR,{"chartType":"column","title":"plotvisonly=true (skip hidden rows when bound to a sheet)",
       "plotvisonly":"true","legend":"none","categories":CATS,"data":D})

# ---------------------------------------------------------------------------
# Slide 6 — Built-in style id + dispBlanksAs
# ---------------------------------------------------------------------------
new_slide("Built-in style & blank handling — style=1..48, dispBlanksAs, dataRange")
ch(TL,{"chartType":"column","style":"2","title":"style=2","legend":"bottom",
       "categories":CATS,"data":D2})
ch(TR,{"chartType":"column","style":"42","title":"style=42","legend":"bottom",
       "categories":CATS,"data":D2})
# dispBlanksAs is Set/Get only — Add first, then Set.
ch(BL,{"chartType":"line","title":"dispBlanksAs=gap (Set after Add)","showMarker":"true",
       "legend":"bottom","categories":CATS,"data":"A:60,90,140,180"})
cli("set",FILE,f"/slide[{slide}]/chart[3]",*P({"dispBlanksAs":"gap"}))
# dataRange is Add-time alternative to data= for sheet-backed sources;
# in a standalone pptx this is largely symbolic — we still demonstrate the syntax,
# then fall back to inline data so the chart renders.
ch(BR,{"chartType":"column","title":"dataRange syntax demo (fallback inline)",
       "dataRange":"Sheet1!A1:D5","legend":"bottom","catTitle":"Quarter",
       "categories":CATS,"data":D2})

# ---------------------------------------------------------------------------
# Slide 7 — chart-axis Set (per-axis post-Add)
# ---------------------------------------------------------------------------
new_slide("chart-axis Set — dispUnits, logBase, minorUnit, visible, labelRotation per-axis")
ch(TL,{"chartType":"column","title":"after: dispUnits=thousands (Set on value axis)",
       "legend":"none","categories":CATS,"data":"Rev:120000,135000,148000,162000"})
cli("set",FILE,f"/slide[{slide}]/chart[1]/axis[@role=value]",
    *P({"dispUnits":"thousands","format":"#,##0","minorUnit":"10000",
        "labelRotation":"0","visible":"true"}))
ch(TR,{"chartType":"line","title":"after: logBase=10 (Set on value axis)",
       "legend":"none","categories":CATS,"data":"A:5,50,500,5000"})
cli("set",FILE,f"/slide[{slide}]/chart[2]/axis[@role=value]",
    *P({"logBase":"10","min":"1","max":"10000","majorGridlines":"true"}))
ch(BL,{"chartType":"column","title":"after: visible=false on value axis",
       "legend":"none","categories":CATS,"data":D})
cli("set",FILE,f"/slide[{slide}]/chart[3]/axis[@role=value]",*P({"visible":"false"}))
ch(BR,{"chartType":"column","title":"after: labelRotation=-45 on category axis",
       "legend":"none","categories":"January,February,March,April","data":D})
cli("set",FILE,f"/slide[{slide}]/chart[4]/axis[@role=category]",
    *P({"labelRotation":"-45","title":"Month","visible":"true"}))

# ---------------------------------------------------------------------------
# Slide 8 — chart-series values=/categories= Set + Get readback round-trip
# ---------------------------------------------------------------------------
new_slide("chart-series mutation — values=, categories= + get-readback round-trip")
ch(TL,{"chartType":"column","title":"before: A=60,90,140,180","legend":"bottom",
       "categories":CATS,"data":D})
# Mutate the values after add
cli("set",FILE,f"/slide[{slide}]/chart[1]/series[1]",*P({"values":"200,150,100,80"}))
note("0.3in","4in","After Set values=200,150,100,80 the series flips downward.")

ch(TR,{"chartType":"column","title":"per-series categories= (range)","legend":"bottom",
       "categories":CATS,"data":D})
# Per-series category override is range-only — note that it requires sheet backing
# so this is a demonstration of the syntax only; effective result depends on workbook.

# Round-trip: change one series, then read it back and stamp the JSON onto the slide
cli("set",FILE,f"/slide[{slide}]/chart[1]/series[1]",
    *P({"name":"Readback Demo","color":"C00000"}))
out = cli("get",FILE,f"/slide[{slide}]/chart[1]/series[1]","--json", capture=True) or ""
# Pretty-print, trim
try:
    obj = json.loads(out)
    if isinstance(obj, dict) and "data" in obj: obj = obj["data"]
    pretty = json.dumps(obj.get("format", obj), indent=2)[:600]
except Exception:
    pretty = out[:600]
cli("add",FILE,f"/slide[{slide}]","--type","shape",
    *P({"text":"chart-series get --json (readback fields alpha/outlineColor/scatterStyle/...):\n"+pretty,
        "size":9,"color":"222222","x":"0.3in","y":"4.25in",
        "width":"6.1in","height":"3in"}))

# chart-axis get-readback — surfaces axisFont/axisMax/axisMin/axisNumFmt/
# axisOrientation/axisTitle/labelOffset/tickLabelSkip read-only fields.
cli("set",FILE,f"/slide[{slide}]/chart[1]/axis[@role=value]",
    *P({"title":"Readback Y","format":"$#,##0","min":"0","max":"300","majorUnit":"75"}))
ax = cli("get",FILE,f"/slide[{slide}]/chart[1]/axis[@role=value]","--json", capture=True) or ""
try:
    obj = json.loads(ax)
    if isinstance(obj, dict) and "data" in obj: obj = obj["data"]
    ax_pretty = json.dumps(obj.get("format", obj), indent=2)[:500]
except Exception:
    ax_pretty = ax[:500]
cli("add",FILE,f"/slide[{slide}]","--type","shape",
    *P({"text":"chart-axis get --json (readback axisFont/axisMax/axisMin/axisNumFmt/axisOrientation/axisTitle/labelOffset/tickLabelSkip):\n"+ax_pretty,
        "size":9,"color":"222222","x":"6.95in","y":"4.25in",
        "width":"6.1in","height":"3in"}))

print(f"Done: {FILE}  ({slide} slides)")
