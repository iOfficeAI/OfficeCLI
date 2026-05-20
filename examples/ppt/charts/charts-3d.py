#!/usr/bin/env python3
"""
3D Charts Showcase — column3d / bar3d / pie3d / line3d / area3d with view3d, gapdepth, shape.

Generates: charts-3d.pptx

  Slide 1  3D families            column3d / bar3d / pie3d / line3d
  Slide 2  area3d & stacked 3D    area3d / stackedColumn3d / percentStackedColumn3d / line3d stacked
  Slide 3  view3d                 different rotX,rotY,perspective angles
  Slide 4  gapdepth               0 / 50 / 150 / 300 (3D bar/column/line/area only)
  Slide 5  bar shapes             box / cylinder / cone / pyramid (bar3d / column3d)
  Slide 6  Title & legend
  Slide 7  Series styling         colors, gradient, transparency, outline, shadow
  Slide 8  Presets

Usage:
  python3 charts-3d.py
"""
import subprocess, os, sys, atexit
FILE = os.path.join(os.path.dirname(__file__), "charts-3d.pptx")
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
CATS="Q1,Q2,Q3,Q4"
D2="East:120,135,148,162;West:95,108,115,128"
D3="East:120,135,148,162;South:95,108,115,128;West:80,90,98,110"
PIE_CATS="North,South,East,West"; PIE_D="Share:30,25,28,17"

if os.path.exists(FILE): os.remove(FILE)
cli("create",FILE); cli("open",FILE)
atexit.register(lambda:(cli("close",FILE),cli("validate",FILE)))

new_slide("3D families — column3d / bar3d / pie3d / line3d")
ch(TL,{"chartType":"column3d","title":"column3d","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"bar3d","title":"bar3d","legend":"bottom","categories":CATS,"data":D2})
ch(BL,{"chartType":"pie3d","title":"pie3d","legend":"right","categories":PIE_CATS,"data":PIE_D})
ch(BR,{"chartType":"line3d","title":"line3d","legend":"bottom","categories":CATS,"data":D2})

new_slide("area3d & stacked 3D")
ch(TL,{"chartType":"area3d","title":"area3d","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"stackedColumn3d","title":"stackedColumn3d","legend":"bottom",
       "categories":CATS,"data":D3})
ch(BL,{"chartType":"percentStackedColumn3d","title":"percentStackedColumn3d","legend":"bottom",
       "categories":CATS,"data":D3})
ch(BR,{"chartType":"stackedBar3d","title":"stackedBar3d","legend":"bottom",
       "categories":CATS,"data":D3})

new_slide("view3d — rotX,rotY,perspective angles")
ch(TL,{"chartType":"column3d","title":"view3d=15,20,30","view3d":"15,20,30",
       "legend":"none","categories":CATS,"data":D2})
ch(TR,{"chartType":"column3d","title":"view3d=30,40,15","view3d":"30,40,15",
       "legend":"none","categories":CATS,"data":D2})
ch(BL,{"chartType":"column3d","title":"view3d=20","view3d":"20",
       "legend":"none","categories":CATS,"data":D2})
ch(BR,{"chartType":"pie3d","title":"pie3d view3d=40,30,30","view3d":"40,30,30",
       "legend":"right","categories":PIE_CATS,"data":PIE_D})

new_slide("gapdepth — 0 / 50 / 150 / 300")
for box,g in zip([TL,TR,BL,BR],[0,50,150,300]):
    ch(box,{"chartType":"column3d","title":f"gapdepth={g}","gapdepth":str(g),
            "legend":"none","categories":CATS,"data":D2})

new_slide("3D bar shapes — box / cylinder / cone / pyramid")
for box,s in zip([TL,TR,BL,BR],["box","cylinder","cone","pyramid"]):
    ch(box,{"chartType":"bar3d","shape":s,"title":f"shape={s}","legend":"none",
            "categories":CATS,"data":D2})

new_slide("Title & legend")
ch(TL,{"chartType":"column3d","title":"Styled title","title.font":"Georgia","title.size":"20",
       "title.color":"4472C4","title.bold":"true","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"column3d","title":"legend=top + legendFont","legend":"top",
       "legendFont":"10:333333:Calibri","categories":CATS,"data":D2})
ch(BL,{"chartType":"column3d","title":"legend.overlay=true","legend":"topRight",
       "legend.overlay":"true","categories":CATS,"data":D2})
ch(BR,{"chartType":"column3d","autotitledeleted":"true","legend":"none","categories":CATS,"data":D2})

new_slide("Series styling — colors, gradient, transparency, outline, shadow")
ch(TL,{"chartType":"column3d","title":"colors + seriesoutline","colors":"4472C4,ED7D31",
       "seriesoutline":"000000:0.5","legend":"bottom","categories":CATS,"data":D2})
ch(TR,{"chartType":"column3d","title":"gradient + seriesshadow",
       "gradient":"FF6600-FFCC00","seriesshadow":"000000-5-45-3-50",
       "legend":"none","categories":CATS,"data":D2})
ch(BL,{"chartType":"column3d","title":"transparency=30","transparency":"30",
       "legend":"bottom","categories":CATS,"data":D2})
ch(BR,{"chartType":"column3d","title":"per-series gradients",
       "gradients":"FF0000-0000FF;00FF00-FFFF00",
       "legend":"bottom","categories":CATS,"data":D2})

new_slide("Presets — preset bundles on 3D charts")
for box,p in zip([TL,TR,BL,BR],["minimal","dark","corporate","colorful"]):
    ch(box,{"chartType":"column3d","preset":p,"title":f"preset={p}",
            "view3d":"15,20,30","legend":"bottom","categories":CATS,"data":D2})

print(f"Done: {FILE}  ({slide} slides)")
