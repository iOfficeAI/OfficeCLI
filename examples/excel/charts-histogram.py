#!/usr/bin/env python3
"""
Histogram Charts — Grand Showcase
==================================

The most thorough, most visually polished histogram demo officecli can
produce. Every binning knob, every styling vocabulary, every canonical
distribution shape, six design themes on one dataset, four font type
specimens, and a cohesive production-grade ML dashboard — all driven by
real copyable officecli CLI commands.

Generates: charts-histogram.xlsx (6 sheets, 29 histograms)

  0-Hero                 1 magazine-grade full-bleed hero poster chart
  1-Binning Lab          6 charts — every binning knob, identical styling
  2-Distribution Zoo     6 canonical real-world distribution shapes
  3-Theme Gallery        6 design themes on the SAME dataset
  4-Typography           4 font-family type specimens
  5-ML Dashboard         6-chart "Production ML Model Report" dashboard

Usage:
  python3 charts-histogram.py
"""

import subprocess, os, atexit, random, math

FILE = "charts-histogram.xlsx"


def cli(cmd):
    """Run: officecli <cmd> — prints stdout/stderr in real time."""
    r = subprocess.run(f"officecli {cmd}", shell=True, capture_output=True, text=True)
    out = (r.stdout or "").strip()
    if out:
        for line in out.split("\n"):
            if line.strip():
                print(f"  {line.strip()}")
    if r.returncode != 0:
        err = (r.stderr or "").strip()
        if err and "UNSUPPORTED" not in err and "process cannot access" not in err:
            print(f"  ERROR: {err}")


# --------------------------------------------------------------------------
# Scaffolding: create file, open it in resident mode (fast subsequent calls),
# and register a graceful close() on exit.
# --------------------------------------------------------------------------
if os.path.exists(FILE):
    os.remove(FILE)

cli(f'create "{FILE}"')
cli(f'open "{FILE}"')
atexit.register(lambda: cli(f'close "{FILE}"'))


# --------------------------------------------------------------------------
# Deterministic sample generators — same seed, same file every regeneration.
# All datasets are CSV-joined once here and reused across sheets.
# --------------------------------------------------------------------------
def csv(values):
    return ",".join(str(v) for v in values)

# The "reference" bell curve — 200 samples around 75±12. Used by the hero,
# the binning lab, the theme gallery, the typography specimens, and the zoo.
random.seed(42)
BELL_200 = sorted(round(random.gauss(75, 12), 1) for _ in range(200))
BELL_CSV = csv(BELL_200)

# Bimodal: two cohorts (beginners ~55, experts ~88) glued together.
random.seed(7)
BIMODAL = sorted(
    [round(random.gauss(55, 6), 1) for _ in range(80)]
    + [round(random.gauss(88, 5), 1) for _ in range(80)]
)
BIMODAL_CSV = csv(BIMODAL)

# Right-skewed / log-normal: classic income shape.
random.seed(11)
LOGNORM = sorted(round(math.exp(random.gauss(3.2, 0.55)), 1) for _ in range(180))
LOGNORM_CSV = csv(LOGNORM)

# Left-skewed: retirement ages — most cluster high, a few retire early.
random.seed(23)
LEFT_SKEW = sorted(round(75 - math.exp(random.gauss(1.6, 0.6)), 1) for _ in range(140))
LEFT_CSV = csv(LEFT_SKEW)

# Uniform: random draws evenly distributed across a range.
random.seed(31)
UNIFORM = sorted(round(random.uniform(0, 100), 1) for _ in range(160))
UNIFORM_CSV = csv(UNIFORM)

# Heavy-tailed (Pareto): most small, tiny fraction catastrophic.
random.seed(47)
PARETO = sorted(round(random.paretovariate(1.6) * 20, 1) for _ in range(200))
PARETO_CSV = csv(PARETO)

# --- ML Dashboard datasets (sheet 5) ---
random.seed(101)
LATENCY_MS = sorted(round(random.paretovariate(1.8) * 15 + 10, 1) for _ in range(250))
LATENCY_CSV = csv(LATENCY_MS)

random.seed(102)
CONFIDENCE = sorted(round(random.betavariate(6, 2) * 100, 2) for _ in range(240))
CONFIDENCE_CSV = csv(CONFIDENCE)

random.seed(103)
ERROR_MAG = sorted(round(abs(random.gauss(0, 1.5)), 3) for _ in range(180))
ERROR_MAG_CSV = csv(ERROR_MAG)

random.seed(104)
TOKEN_LEN = sorted(
    [max(1, round(random.gauss(180, 40))) for _ in range(100)]
    + [max(1, round(random.gauss(520, 90))) for _ in range(80)]
)
TOKEN_CSV = csv(TOKEN_LEN)

random.seed(105)
GPU_UTIL = sorted(round(min(99.0, max(30.0, random.gauss(82, 8))), 1) for _ in range(200))
GPU_CSV = csv(GPU_UTIL)

random.seed(106)
COST_REQ = sorted(round(math.exp(random.gauss(-3.2, 0.9)) * 1000, 3) for _ in range(220))
COST_CSV = csv(COST_REQ)


# ==========================================================================
# Sheet 0: "0-Hero" — the full-bleed magazine hero poster
#
# A single giant chart using EVERY histogram knob at once:
#   - Dark "Midnight Academia" palette: navy plot area, gold bars, cream text
#   - title.*  (color/size/bold/font/shadow)
#   - series.shadow + fill
#   - axisline + axisfont + axisTitle.*
#   - plotareafill / plotarea.border / chartareafill / chartarea.border
#   - axismin / axismax / majorunit (locked Y scale)
#   - gridlineColor
#   - dataLabels + datalabels.numfmt
#   - legend=top + legend.overlay + legendfont
#   - intervalClosed=l + explicit binCount
#
# This chart is the "representative sample" — if it renders correctly, the
# entire histogram pipeline is healthy.
# ==========================================================================
print("\n--- 0-Hero ---")
cli(f'set "{FILE}" /Sheet1 --prop name="0-Hero"')

# officecli add charts-histogram.xlsx "/0-Hero" --type chart \
#   --prop chartType=histogram \
#   --prop title="The Shape of Data · 200-sample bell curve" \
#   --prop title.color=F5F1E0 --prop title.size=22 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop "title.shadow=000000-8-45-4-70" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=24 --prop intervalClosed=l \
#   --prop fill=F0C96A --prop "series.shadow=000000-8-45-4-60" \
#   --prop axismin=0 --prop axismax=28 --prop majorunit=4 \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Frequency" \
#   --prop axisTitle.color=C9B87A --prop axisTitle.size=13 \
#   --prop axisTitle.bold=true --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=10:B8B090:Helvetica Neue" \
#   --prop "axisline=6A6448:1.5" \
#   --prop gridlineColor=2F3544 \
#   --prop plotareafill=1A1F2C --prop "plotarea.border=3A3E4E:1.25" \
#   --prop chartareafill=0B0F18 --prop "chartarea.border=2A2E3E:1" \
#   --prop dataLabels=true --prop "datalabels.numfmt=0" \
#   --prop legend=top --prop legend.overlay=false \
#   --prop "legendfont=11:D4C994:Helvetica Neue" \
#   --prop x=0 --prop y=0 --prop width=27 --prop height=38
# Features: EVERY knob — title/series/axis/plotarea/chartarea/shadow/scaling/legend/datalabel
cli(f'add "{FILE}" "/0-Hero" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="The Shape of Data · 200-sample bell curve"'
    f' --prop title.color=F5F1E0 --prop title.size=22 --prop title.bold=true'
    f' --prop title.font="Helvetica Neue"'
    f' --prop "title.shadow=000000-8-45-4-70"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=24 --prop intervalClosed=l'
    f' --prop fill=F0C96A --prop "series.shadow=000000-8-45-4-60"'
    f' --prop axismin=0 --prop axismax=28 --prop majorunit=4'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Frequency"'
    f' --prop axisTitle.color=C9B87A --prop axisTitle.size=13'
    f' --prop axisTitle.bold=true --prop axisTitle.font="Helvetica Neue"'
    f' --prop "axisfont=10:B8B090:Helvetica Neue"'
    f' --prop "axisline=6A6448:1.5"'
    f' --prop gridlineColor=2F3544'
    f' --prop plotareafill=1A1F2C --prop "plotarea.border=3A3E4E:1.25"'
    f' --prop chartareafill=0B0F18 --prop "chartarea.border=2A2E3E:1"'
    f' --prop dataLabels=true --prop "datalabels.numfmt=0"'
    f' --prop legend=top --prop legend.overlay=false'
    f' --prop "legendfont=11:D4C994:Helvetica Neue"'
    f' --prop x=0 --prop y=0 --prop width=27 --prop height=38')


# ==========================================================================
# Sheet 1: "1-Binning Lab"
#
# Six histograms, SAME dataset (BELL_200), IDENTICAL typography / colors /
# frames — the ONLY thing that varies is the binning strategy. Put side by
# side, this sheet is the "Rosetta stone": once you see how each binning
# knob reshapes the bars, you'll never be confused about which to use.
#
#   ┌──────────┬──────────┐
#   │ 1. auto  │ 2. count │
#   ├──────────┼──────────┤
#   │ 3. fine  │ 4. width │
#   ├──────────┼──────────┤
#   │ 5. fence │ 6. lclos │
#   └──────────┴──────────┘
# ==========================================================================
print("\n--- 1-Binning Lab ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Binning Lab"')

# Shared "clean lab" style — every chart on this sheet wears the exact same
# outfit so the bin-shape difference is the only visible variable.
LAB = (
    ' --prop fill=4472C4'
    ' --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true'
    ' --prop title.font="Helvetica Neue"'
    ' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    ' --prop axisTitle.color=6B7280 --prop axisTitle.size=10'
    ' --prop axisTitle.font="Helvetica Neue"'
    ' --prop "axisfont=9:6B7280:Helvetica Neue"'
    ' --prop gridlineColor=F0F0F0'
    ' --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75"'
    ' --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75"'
    ' --prop "axisline=9CA3AF:0.75"'
)

# officecli add charts-histogram.xlsx "/1-Binning Lab" --type chart \
#   --prop chartType=histogram \
#   --prop title="1 · Auto-binning (Excel default)" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop fill=4472C4 \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=0 --prop y=0 --prop width=13 --prop height=18
# Features: no binCount, no binSize — Excel picks the bin count automatically.
cli(f'add "{FILE}" "/1-Binning Lab" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="1 · Auto-binning (Excel default)"'
    f' --prop series1=Samples:{BELL_CSV}'
    f'{LAB}'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/1-Binning Lab" --type chart \
#   --prop chartType=histogram \
#   --prop title="2 · binCount=8 (coarse)" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=8 \
#   --prop fill=4472C4 \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=14 --prop y=0 --prop width=13 --prop height=18
# Features: binCount=8 — coarse. Fewer, wider bars. Good for "what's the mode?"
cli(f'add "{FILE}" "/1-Binning Lab" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="2 · binCount=8 (coarse)"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=8'
    f'{LAB}'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/1-Binning Lab" --type chart \
#   --prop chartType=histogram \
#   --prop title="3 · binCount=32 (fine)" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=32 \
#   --prop fill=4472C4 \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=0 --prop y=19 --prop width=13 --prop height=18
# Features: binCount=32 — fine. Many narrow bars. Good for "is it really Gaussian?"
cli(f'add "{FILE}" "/1-Binning Lab" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="3 · binCount=32 (fine)"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=32'
    f'{LAB}'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/1-Binning Lab" --type chart \
#   --prop chartType=histogram \
#   --prop title="4 · binSize=5 (fixed-width bins)" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binSize=5 \
#   --prop fill=4472C4 \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=14 --prop y=19 --prop width=13 --prop height=18
# Features: binSize=5 — fixed bin width. Use when you want human-friendly
# bin boundaries (multiples of 5, 10, etc) regardless of data range.
cli(f'add "{FILE}" "/1-Binning Lab" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="4 · binSize=5 (fixed-width bins)"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binSize=5'
    f'{LAB}'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/1-Binning Lab" --type chart \
#   --prop chartType=histogram \
#   --prop title="5 · underflow=55 · overflow=95 (fencing)" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binSize=5 --prop underflowBin=55 --prop overflowBin=95 \
#   --prop fill=4472C4 \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=0 --prop y=38 --prop width=13 --prop height=18
# Features: underflowBin=55 + overflowBin=95 — outlier fencing. Everything
# below 55 or above 95 collapses into a single <55 / >95 bar.
cli(f'add "{FILE}" "/1-Binning Lab" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="5 · underflow=55 · overflow=95 (fencing)"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binSize=5 --prop underflowBin=55 --prop overflowBin=95'
    f'{LAB}'
    f' --prop x=0 --prop y=38 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/1-Binning Lab" --type chart \
#   --prop chartType=histogram \
#   --prop title="6 · [a,b) intervals + gapWidth=30" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=16 --prop intervalClosed=l --prop gapWidth=30 \
#   --prop fill=4472C4 \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=14 --prop y=38 --prop width=13 --prop height=18
# Features: intervalClosed=l (half-open [a,b)) + gapWidth=30 — shows the
# "left-closed" variant AND pushes bars apart so you can see each one.
# Useful when the dataset has values lying exactly on a bin boundary.
cli(f'add "{FILE}" "/1-Binning Lab" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="6 · [a,b) intervals + gapWidth=30"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=16 --prop intervalClosed=l --prop gapWidth=30'
    f'{LAB}'
    f' --prop x=14 --prop y=38 --prop width=13 --prop height=18')


# ==========================================================================
# Sheet 2: "2-Distribution Zoo"
#
# A cohesive 2x3 gallery of the canonical distribution shapes you'll see
# in production data. Pattern recognition: if you ever see one of these
# shapes in a telemetry chart, you know immediately what's going on.
#
# Every chart shares the same typography + plot/chart area frames; only
# the fill color and data change. Uses different binning strategies
# appropriate to each distribution.
# ==========================================================================
print("\n--- 2-Distribution Zoo ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Distribution Zoo"')

ZOO = (
    ' --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true'
    ' --prop title.font="Helvetica Neue"'
    ' --prop axisTitle.color=6B7280 --prop axisTitle.size=10'
    ' --prop axisTitle.font="Helvetica Neue"'
    ' --prop "axisfont=9:6B7280:Helvetica Neue"'
    ' --prop gridlineColor=EFEFEF'
    ' --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75"'
    ' --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75"'
    ' --prop "axisline=9CA3AF:0.75"'
)

# officecli add charts-histogram.xlsx "/2-Distribution Zoo" --type chart \
#   --prop chartType=histogram \
#   --prop title="Normal · bell curve (reference)" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=2F5597 \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=EFEFEF \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=0 --prop y=0 --prop width=13 --prop height=18
# Features: classic bell curve reference, binCount=18, midnight blue fill.
cli(f'add "{FILE}" "/2-Distribution Zoo" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Normal · bell curve (reference)"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=2F5597'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f'{ZOO}'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/2-Distribution Zoo" --type chart \
#   --prop chartType=histogram \
#   --prop title="Bimodal · two hidden cohorts" \
#   --prop series1="Score:<160 bimodal values>" \
#   --prop binCount=22 --prop fill=ED7D31 \
#   --prop xAxisTitle="Test score" --prop yAxisTitle="Students" \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=EFEFEF \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=14 --prop y=0 --prop width=13 --prop height=18
# Features: bimodal — two hidden populations. Narrow bins reveal the split.
cli(f'add "{FILE}" "/2-Distribution Zoo" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Bimodal · two hidden cohorts"'
    f' --prop series1=Score:{BIMODAL_CSV}'
    f' --prop binCount=22 --prop fill=ED7D31'
    f' --prop xAxisTitle="Test score" --prop yAxisTitle="Students"'
    f'{ZOO}'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/2-Distribution Zoo" --type chart \
#   --prop chartType=histogram \
#   --prop title="Right-skewed · log-normal (income)" \
#   --prop series1="Income:<180 log-normal values>" \
#   --prop binCount=20 --prop fill=70AD47 \
#   --prop xAxisTitle="Monthly income ($k)" --prop yAxisTitle="People" \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=EFEFEF \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=0 --prop y=19 --prop width=13 --prop height=18
# Features: right-skewed log-normal. Mean >> median, long tail to the right.
cli(f'add "{FILE}" "/2-Distribution Zoo" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Right-skewed · log-normal (income)"'
    f' --prop series1=Income:{LOGNORM_CSV}'
    f' --prop binCount=20 --prop fill=70AD47'
    f' --prop xAxisTitle="Monthly income ($k)" --prop yAxisTitle="People"'
    f'{ZOO}'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/2-Distribution Zoo" --type chart \
#   --prop chartType=histogram \
#   --prop title="Left-skewed · retirement ages" \
#   --prop series1="Age:<140 left-skewed values>" \
#   --prop binCount=18 --prop fill=7030A0 \
#   --prop xAxisTitle="Age at retirement" --prop yAxisTitle="Retirees" \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=EFEFEF \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=14 --prop y=19 --prop width=13 --prop height=18
# Features: left-skewed — retirement ages cluster high, tail stretches left.
cli(f'add "{FILE}" "/2-Distribution Zoo" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Left-skewed · retirement ages"'
    f' --prop series1=Age:{LEFT_CSV}'
    f' --prop binCount=18 --prop fill=7030A0'
    f' --prop xAxisTitle="Age at retirement" --prop yAxisTitle="Retirees"'
    f'{ZOO}'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/2-Distribution Zoo" --type chart \
#   --prop chartType=histogram \
#   --prop title="Uniform · flat floor" \
#   --prop series1="Draws:<160 uniform values>" \
#   --prop binSize=10 --prop fill=00B0F0 \
#   --prop xAxisTitle="Random draw (0-100)" --prop yAxisTitle="Count" \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=EFEFEF \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=0 --prop y=38 --prop width=13 --prop height=18
# Features: uniform — every value equally likely. binSize emphasizes the
# "flat floor" visual tell.
cli(f'add "{FILE}" "/2-Distribution Zoo" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Uniform · flat floor"'
    f' --prop series1=Draws:{UNIFORM_CSV}'
    f' --prop binSize=10 --prop fill=00B0F0'
    f' --prop xAxisTitle="Random draw (0-100)" --prop yAxisTitle="Count"'
    f'{ZOO}'
    f' --prop x=0 --prop y=38 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/2-Distribution Zoo" --type chart \
#   --prop chartType=histogram \
#   --prop title="Heavy-tailed · Pareto (overflow=250)" \
#   --prop series1="Latency:<200 Pareto values>" \
#   --prop binSize=20 --prop overflowBin=250 --prop fill=C00000 \
#   --prop xAxisTitle="Latency (ms)" --prop yAxisTitle="Requests" \
#   --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=9:6B7280:Helvetica Neue" \
#   --prop gridlineColor=EFEFEF \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=14 --prop y=38 --prop width=13 --prop height=18
# Features: heavy-tailed Pareto + overflowBin. Fences the catastrophic tail
# so the interesting bulk of the distribution stays readable.
cli(f'add "{FILE}" "/2-Distribution Zoo" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Heavy-tailed · Pareto (overflow=250)"'
    f' --prop series1=Latency:{PARETO_CSV}'
    f' --prop binSize=20 --prop overflowBin=250 --prop fill=C00000'
    f' --prop xAxisTitle="Latency (ms)" --prop yAxisTitle="Requests"'
    f'{ZOO}'
    f' --prop x=14 --prop y=38 --prop width=13 --prop height=18')


# ==========================================================================
# Sheet 3: "3-Theme Gallery"
#
# Six complete design themes applied to the SAME bell-curve dataset. Each
# theme is a coordinated palette: plot-area fill, chart-area fill, series
# fill, gridline color, axis line color, tick-label color, title color,
# title font — all chosen to read as one coherent mood.
#
# Grid:
#   ┌─────────────┬─────────────┐
#   │ 1. Midnight │ 2. Sunset   │
#   ├─────────────┼─────────────┤
#   │ 3. Forest   │ 4. Mono     │
#   ├─────────────┼─────────────┤
#   │ 5. Neon     │ 6. Pastel   │
#   └─────────────┴─────────────┘
# ==========================================================================
print("\n--- 3-Theme Gallery ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Theme Gallery"')

# officecli add charts-histogram.xlsx "/3-Theme Gallery" --type chart \
#   --prop chartType=histogram \
#   --prop title="Midnight Academia" \
#   --prop title.color=F5F1E0 --prop title.size=14 --prop title.bold=true \
#   --prop title.font="Georgia" \
#   --prop "title.shadow=000000-6-45-3-70" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=F0C96A \
#   --prop "series.shadow=000000-6-45-3-55" \
#   --prop plotareafill=1A1F2C --prop "plotarea.border=3A3E4E:1" \
#   --prop chartareafill=0B0F18 --prop "chartarea.border=2A2E3E:0.75" \
#   --prop gridlineColor=2F3544 \
#   --prop "axisfont=9:B8B090:Georgia" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=C9B87A --prop axisTitle.size=10 \
#   --prop axisTitle.font="Georgia" \
#   --prop "axisline=5A5848:1" \
#   --prop x=0 --prop y=0 --prop width=13 --prop height=18
# Features: dark plot area, gold bars, series.shadow, title.shadow
cli(f'add "{FILE}" "/3-Theme Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Midnight Academia"'
    f' --prop title.color=F5F1E0 --prop title.size=14 --prop title.bold=true'
    f' --prop title.font="Georgia"'
    f' --prop "title.shadow=000000-6-45-3-70"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=F0C96A'
    f' --prop "series.shadow=000000-6-45-3-55"'
    f' --prop plotareafill=1A1F2C --prop "plotarea.border=3A3E4E:1"'
    f' --prop chartareafill=0B0F18 --prop "chartarea.border=2A2E3E:0.75"'
    f' --prop gridlineColor=2F3544'
    f' --prop "axisfont=9:B8B090:Georgia"'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=C9B87A --prop axisTitle.size=10'
    f' --prop axisTitle.font="Georgia"'
    f' --prop "axisline=5A5848:1"'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/3-Theme Gallery" --type chart \
#   --prop chartType=histogram \
#   --prop title="Sunset Terracotta" \
#   --prop title.color=3F2818 --prop title.size=14 --prop title.bold=true \
#   --prop title.font="Georgia" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=E85D4A \
#   --prop plotareafill=FFF5E8 --prop "plotarea.border=F0D8B0:1" \
#   --prop chartareafill=FFE6C7 --prop "chartarea.border=E6BC88:1" \
#   --prop gridlineColor=F5C98A \
#   --prop "axisfont=9:6B4A2A:Georgia" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=A8522C --prop axisTitle.size=10 \
#   --prop axisTitle.font="Georgia" \
#   --prop "axisline=C08050:1" \
#   --prop x=14 --prop y=0 --prop width=13 --prop height=18
# Theme 2 · Sunset Terracotta (warm cream + coral, serif)
cli(f'add "{FILE}" "/3-Theme Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Sunset Terracotta"'
    f' --prop title.color=3F2818 --prop title.size=14 --prop title.bold=true'
    f' --prop title.font="Georgia"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=E85D4A'
    f' --prop plotareafill=FFF5E8 --prop "plotarea.border=F0D8B0:1"'
    f' --prop chartareafill=FFE6C7 --prop "chartarea.border=E6BC88:1"'
    f' --prop gridlineColor=F5C98A'
    f' --prop "axisfont=9:6B4A2A:Georgia"'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=A8522C --prop axisTitle.size=10'
    f' --prop axisTitle.font="Georgia"'
    f' --prop "axisline=C08050:1"'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/3-Theme Gallery" --type chart \
#   --prop chartType=histogram \
#   --prop title="Forest Parchment" \
#   --prop title.color=1F3A1F --prop title.size=14 --prop title.bold=true \
#   --prop title.font="Georgia" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=2F5D3A \
#   --prop plotareafill=F3EDD8 --prop "plotarea.border=C8B890:1" \
#   --prop chartareafill=EADFBE --prop "chartarea.border=A89858:1" \
#   --prop gridlineColor=C0B888 \
#   --prop "axisfont=9:4A5A3A:Georgia" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=3F5A2F --prop axisTitle.size=10 \
#   --prop axisTitle.font="Georgia" \
#   --prop "axisline=6A7A4A:1" \
#   --prop x=0 --prop y=19 --prop width=13 --prop height=18
# Theme 3 · Forest Parchment (beige + forest green, serif)
cli(f'add "{FILE}" "/3-Theme Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Forest Parchment"'
    f' --prop title.color=1F3A1F --prop title.size=14 --prop title.bold=true'
    f' --prop title.font="Georgia"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=2F5D3A'
    f' --prop plotareafill=F3EDD8 --prop "plotarea.border=C8B890:1"'
    f' --prop chartareafill=EADFBE --prop "chartarea.border=A89858:1"'
    f' --prop gridlineColor=C0B888'
    f' --prop "axisfont=9:4A5A3A:Georgia"'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=3F5A2F --prop axisTitle.size=10'
    f' --prop axisTitle.font="Georgia"'
    f' --prop "axisline=6A7A4A:1"'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/3-Theme Gallery" --type chart \
#   --prop chartType=histogram \
#   --prop title="Editorial Mono" \
#   --prop title.color=111111 --prop title.size=14 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=2A2A2A \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=CCCCCC:0.75" \
#   --prop chartareafill=FAFAFA --prop "chartarea.border=E0E0E0:0.75" \
#   --prop gridlineColor=EEEEEE \
#   --prop "axisfont=9:555555:Helvetica Neue" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=333333 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisline=888888:1" \
#   --prop x=14 --prop y=19 --prop width=13 --prop height=18
# Theme 4 · Editorial Mono (pure grayscale, sans)
cli(f'add "{FILE}" "/3-Theme Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Editorial Mono"'
    f' --prop title.color=111111 --prop title.size=14 --prop title.bold=true'
    f' --prop title.font="Helvetica Neue"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=2A2A2A'
    f' --prop plotareafill=FFFFFF --prop "plotarea.border=CCCCCC:0.75"'
    f' --prop chartareafill=FAFAFA --prop "chartarea.border=E0E0E0:0.75"'
    f' --prop gridlineColor=EEEEEE'
    f' --prop "axisfont=9:555555:Helvetica Neue"'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=333333 --prop axisTitle.size=10'
    f' --prop axisTitle.font="Helvetica Neue"'
    f' --prop "axisline=888888:1"'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/3-Theme Gallery" --type chart \
#   --prop chartType=histogram \
#   --prop title="Neon Terminal" \
#   --prop title.color=00F0C8 --prop title.size=14 --prop title.bold=true \
#   --prop title.font="Courier New" \
#   --prop "title.shadow=00F0C8-6-45-0-40" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=00F0C8 \
#   --prop "series.shadow=00F0C8-8-45-0-45" \
#   --prop plotareafill=0A0A14 --prop "plotarea.border=1F2F3F:1" \
#   --prop chartareafill=000008 --prop "chartarea.border=1F1F2F:1" \
#   --prop gridlineColor=1A2A3A \
#   --prop "axisfont=9:00D0E8:Courier New" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=00D0E8 --prop axisTitle.size=10 \
#   --prop axisTitle.font="Courier New" \
#   --prop "axisline=00707F:1" \
#   --prop x=0 --prop y=38 --prop width=13 --prop height=18
# Theme 5 · Neon Terminal (black + electric cyan, mono)
cli(f'add "{FILE}" "/3-Theme Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Neon Terminal"'
    f' --prop title.color=00F0C8 --prop title.size=14 --prop title.bold=true'
    f' --prop title.font="Courier New"'
    f' --prop "title.shadow=00F0C8-6-45-0-40"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=00F0C8'
    f' --prop "series.shadow=00F0C8-8-45-0-45"'
    f' --prop plotareafill=0A0A14 --prop "plotarea.border=1F2F3F:1"'
    f' --prop chartareafill=000008 --prop "chartarea.border=1F1F2F:1"'
    f' --prop gridlineColor=1A2A3A'
    f' --prop "axisfont=9:00D0E8:Courier New"'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=00D0E8 --prop axisTitle.size=10'
    f' --prop axisTitle.font="Courier New"'
    f' --prop "axisline=00707F:1"'
    f' --prop x=0 --prop y=38 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/3-Theme Gallery" --type chart \
#   --prop chartType=histogram \
#   --prop title="Pastel Bloom" \
#   --prop title.color=5A3C4A --prop title.size=14 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=F5A7C8 \
#   --prop plotareafill=FDF4F8 --prop "plotarea.border=F0D0E0:1" \
#   --prop chartareafill=FAEDF2 --prop "chartarea.border=F0C0D8:1" \
#   --prop gridlineColor=F5D8E5 \
#   --prop "axisfont=9:8A6878:Helvetica Neue" \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=A04C6A --prop axisTitle.size=10 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisline=C888A0:1" \
#   --prop x=14 --prop y=38 --prop width=13 --prop height=18
# Theme 6 · Pastel Bloom (lavender cream + rose, sans)
cli(f'add "{FILE}" "/3-Theme Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Pastel Bloom"'
    f' --prop title.color=5A3C4A --prop title.size=14 --prop title.bold=true'
    f' --prop title.font="Helvetica Neue"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=F5A7C8'
    f' --prop plotareafill=FDF4F8 --prop "plotarea.border=F0D0E0:1"'
    f' --prop chartareafill=FAEDF2 --prop "chartarea.border=F0C0D8:1"'
    f' --prop gridlineColor=F5D8E5'
    f' --prop "axisfont=9:8A6878:Helvetica Neue"'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=A04C6A --prop axisTitle.size=10'
    f' --prop axisTitle.font="Helvetica Neue"'
    f' --prop "axisline=C888A0:1"'
    f' --prop x=14 --prop y=38 --prop width=13 --prop height=18')


# ==========================================================================
# Sheet 4: "4-Typography"
#
# Four font-family "type specimens". Same data, same geometry, same colors —
# only the font varies. Side-by-side, this shows how typography alone reads
# as tone: Helvetica is corporate, Georgia is editorial, Courier is data,
# Verdana is approachable.
# ==========================================================================
print("\n--- 4-Typography ---")
cli(f'add "{FILE}" / --type sheet --prop name="4-Typography"')

# officecli add charts-histogram.xlsx "/4-Typography" --type chart \
#   --prop chartType=histogram \
#   --prop title="Helvetica Neue · modern sans" \
#   --prop title.color=1F2937 --prop title.size=16 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=4472C4 \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=4472C4 --prop axisTitle.size=11 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=10:6B7280:Helvetica Neue" \
#   --prop gridlineColor=EEEEEE \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop x=0 --prop y=0 --prop width=13 --prop height=18
# Specimen 1 · Helvetica Neue (modern sans — dashboards, corporate reports)
cli(f'add "{FILE}" "/4-Typography" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Helvetica Neue · modern sans"'
    f' --prop title.color=1F2937 --prop title.size=16 --prop title.bold=true'
    f' --prop title.font="Helvetica Neue"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=4472C4'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=4472C4 --prop axisTitle.size=11'
    f' --prop axisTitle.font="Helvetica Neue"'
    f' --prop "axisfont=10:6B7280:Helvetica Neue"'
    f' --prop gridlineColor=EEEEEE'
    f' --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75"'
    f' --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75"'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/4-Typography" --type chart \
#   --prop chartType=histogram \
#   --prop title="Georgia · editorial serif" \
#   --prop title.color=3F2818 --prop title.size=16 --prop title.bold=true \
#   --prop title.font="Georgia" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=A8522C \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=A8522C --prop axisTitle.size=11 \
#   --prop axisTitle.font="Georgia" \
#   --prop "axisfont=10:6B4A2A:Georgia" \
#   --prop gridlineColor=F0E8D8 \
#   --prop plotareafill=FFFBF3 --prop "plotarea.border=E8D8B8:0.75" \
#   --prop chartareafill=FDF6E8 --prop "chartarea.border=E8D8B8:0.75" \
#   --prop x=14 --prop y=0 --prop width=13 --prop height=18
# Specimen 2 · Georgia (editorial serif — magazines, long-form reports)
cli(f'add "{FILE}" "/4-Typography" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Georgia · editorial serif"'
    f' --prop title.color=3F2818 --prop title.size=16 --prop title.bold=true'
    f' --prop title.font="Georgia"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=A8522C'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=A8522C --prop axisTitle.size=11'
    f' --prop axisTitle.font="Georgia"'
    f' --prop "axisfont=10:6B4A2A:Georgia"'
    f' --prop gridlineColor=F0E8D8'
    f' --prop plotareafill=FFFBF3 --prop "plotarea.border=E8D8B8:0.75"'
    f' --prop chartareafill=FDF6E8 --prop "chartarea.border=E8D8B8:0.75"'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/4-Typography" --type chart \
#   --prop chartType=histogram \
#   --prop title="Courier New · data mono" \
#   --prop title.color=1A3A1A --prop title.size=16 --prop title.bold=true \
#   --prop title.font="Courier New" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=2F8F4F \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=2F8F4F --prop axisTitle.size=11 \
#   --prop axisTitle.font="Courier New" \
#   --prop "axisfont=10:3A5A3A:Courier New" \
#   --prop gridlineColor=E0EDE0 \
#   --prop plotareafill=F7FBF7 --prop "plotarea.border=C8DCC8:0.75" \
#   --prop chartareafill=F0F7F0 --prop "chartarea.border=C8DCC8:0.75" \
#   --prop x=0 --prop y=19 --prop width=13 --prop height=18
# Specimen 3 · Courier New (monospace — data, telemetry, engineering)
cli(f'add "{FILE}" "/4-Typography" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Courier New · data mono"'
    f' --prop title.color=1A3A1A --prop title.size=16 --prop title.bold=true'
    f' --prop title.font="Courier New"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=2F8F4F'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=2F8F4F --prop axisTitle.size=11'
    f' --prop axisTitle.font="Courier New"'
    f' --prop "axisfont=10:3A5A3A:Courier New"'
    f' --prop gridlineColor=E0EDE0'
    f' --prop plotareafill=F7FBF7 --prop "plotarea.border=C8DCC8:0.75"'
    f' --prop chartareafill=F0F7F0 --prop "chartarea.border=C8DCC8:0.75"'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/4-Typography" --type chart \
#   --prop chartType=histogram \
#   --prop title="Verdana · friendly sans" \
#   --prop title.color=4A2B6A --prop title.size=16 --prop title.bold=true \
#   --prop title.font="Verdana" \
#   --prop series1="Samples:<200 bell values>" \
#   --prop binCount=18 --prop fill=8E4DBB \
#   --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
#   --prop axisTitle.color=8E4DBB --prop axisTitle.size=11 \
#   --prop axisTitle.font="Verdana" \
#   --prop "axisfont=10:6B4A8A:Verdana" \
#   --prop gridlineColor=ECE0F4 \
#   --prop plotareafill=FCF7FF --prop "plotarea.border=D8C4E8:0.75" \
#   --prop chartareafill=F6EDFA --prop "chartarea.border=D8C4E8:0.75" \
#   --prop x=14 --prop y=19 --prop width=13 --prop height=18
# Specimen 4 · Verdana (friendly sans — onboarding, public-facing UI)
cli(f'add "{FILE}" "/4-Typography" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Verdana · friendly sans"'
    f' --prop title.color=4A2B6A --prop title.size=16 --prop title.bold=true'
    f' --prop title.font="Verdana"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18 --prop fill=8E4DBB'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=8E4DBB --prop axisTitle.size=11'
    f' --prop axisTitle.font="Verdana"'
    f' --prop "axisfont=10:6B4A8A:Verdana"'
    f' --prop gridlineColor=ECE0F4'
    f' --prop plotareafill=FCF7FF --prop "plotarea.border=D8C4E8:0.75"'
    f' --prop chartareafill=F6EDFA --prop "chartarea.border=D8C4E8:0.75"'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')


# ==========================================================================
# Sheet 5: "5-ML Dashboard"
#
# A cohesive six-chart "Production ML Model Report". Every chart wears the
# same corporate dashboard uniform — same typography, same frames, same
# gridlines — but each shows a different slice of the model's behavior,
# deliberately using a different color + binning strategy so the six read
# as a single dashboard at a glance.
#
#   Row 1:  Inference latency (ms)   |  Prediction confidence (%)
#   Row 2:  |Residual| (logit)       |  Token length (chars)
#   Row 3:  GPU utilization (%)      |  Cost per request ($ × 0.001)
# ==========================================================================
print("\n--- 5-ML Dashboard ---")
cli(f'add "{FILE}" / --type sheet --prop name="5-ML Dashboard"')

DASH = (
    ' --prop title.color=1F2937 --prop title.size=12 --prop title.bold=true'
    ' --prop title.font="Helvetica Neue"'
    ' --prop axisTitle.color=6B7280 --prop axisTitle.size=9'
    ' --prop axisTitle.font="Helvetica Neue"'
    ' --prop "axisfont=8:6B7280:Helvetica Neue"'
    ' --prop gridlineColor=F0F0F0'
    ' --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75"'
    ' --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75"'
    ' --prop "axisline=9CA3AF:0.75"'
    ' --prop dataLabels=false'
)

# officecli add charts-histogram.xlsx "/5-ML Dashboard" --type chart \
#   --prop chartType=histogram \
#   --prop title="Inference Latency · p50-p99 (ms)" \
#   --prop series1="Latency:<250 Pareto latency values>" \
#   --prop binSize=25 --prop overflowBin=300 --prop fill=EF4444 \
#   --prop "series.shadow=EF4444-4-45-2-25" \
#   --prop xAxisTitle="Latency (ms)" --prop yAxisTitle="Requests" \
#   --prop title.color=1F2937 --prop title.size=12 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=9 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=8:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop dataLabels=false \
#   --prop x=0 --prop y=0 --prop width=13 --prop height=18
# 1 · Inference Latency — heavy-tail, overflow-fenced, red for "watch this"
cli(f'add "{FILE}" "/5-ML Dashboard" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Inference Latency · p50-p99 (ms)"'
    f' --prop series1=Latency:{LATENCY_CSV}'
    f' --prop binSize=25 --prop overflowBin=300 --prop fill=EF4444'
    f' --prop "series.shadow=EF4444-4-45-2-25"'
    f' --prop xAxisTitle="Latency (ms)" --prop yAxisTitle="Requests"'
    f'{DASH}'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/5-ML Dashboard" --type chart \
#   --prop chartType=histogram \
#   --prop title="Prediction Confidence" \
#   --prop series1="Confidence:<240 beta confidence values>" \
#   --prop binSize=5 --prop fill=10B981 \
#   --prop axismin=0 --prop majorunit=50 \
#   --prop xAxisTitle="Softmax confidence (%)" --prop yAxisTitle="Samples" \
#   --prop title.color=1F2937 --prop title.size=12 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=9 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=8:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop dataLabels=false \
#   --prop x=14 --prop y=0 --prop width=13 --prop height=18
# 2 · Prediction Confidence — beta-like, axismin/max locked to 0..100
cli(f'add "{FILE}" "/5-ML Dashboard" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Prediction Confidence"'
    f' --prop series1=Confidence:{CONFIDENCE_CSV}'
    f' --prop binSize=5 --prop fill=10B981'
    f' --prop axismin=0 --prop majorunit=50'
    f' --prop xAxisTitle="Softmax confidence (%)" --prop yAxisTitle="Samples"'
    f'{DASH}'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/5-ML Dashboard" --type chart \
#   --prop chartType=histogram \
#   --prop title="|Residual| · model calibration" \
#   --prop series1="Residual:<180 half-normal error values>" \
#   --prop binSize=0.25 --prop intervalClosed=l --prop fill=F59E0B \
#   --prop xAxisTitle="|y - ŷ| (logit)" --prop yAxisTitle="Samples" \
#   --prop title.color=1F2937 --prop title.size=12 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=9 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=8:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop dataLabels=false \
#   --prop x=0 --prop y=19 --prop width=13 --prop height=18
# 3 · Residual Magnitude — half-normal, intervalClosed=l so bin=0 catches zeros
cli(f'add "{FILE}" "/5-ML Dashboard" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="|Residual| · model calibration"'
    f' --prop series1=Residual:{ERROR_MAG_CSV}'
    f' --prop binSize=0.25 --prop intervalClosed=l --prop fill=F59E0B'
    f' --prop xAxisTitle="|y - ŷ| (logit)" --prop yAxisTitle="Samples"'
    f'{DASH}'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/5-ML Dashboard" --type chart \
#   --prop chartType=histogram \
#   --prop title="Token Length · short vs long prompts" \
#   --prop series1="Tokens:<180 bimodal token-length values>" \
#   --prop binCount=24 --prop fill=6366F1 \
#   --prop xAxisTitle="Tokens" --prop yAxisTitle="Requests" \
#   --prop title.color=1F2937 --prop title.size=12 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=9 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=8:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop dataLabels=false \
#   --prop x=14 --prop y=19 --prop width=13 --prop height=18
# 4 · Token Length — bimodal (short prompts vs long prompts)
cli(f'add "{FILE}" "/5-ML Dashboard" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Token Length · short vs long prompts"'
    f' --prop series1=Tokens:{TOKEN_CSV}'
    f' --prop binCount=24 --prop fill=6366F1'
    f' --prop xAxisTitle="Tokens" --prop yAxisTitle="Requests"'
    f'{DASH}'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/5-ML Dashboard" --type chart \
#   --prop chartType=histogram \
#   --prop title="GPU Utilization" \
#   --prop series1="GPU:<200 normal GPU utilization values>" \
#   --prop binSize=5 --prop fill=8B5CF6 \
#   --prop axismin=0 --prop axismax=50 --prop majorunit=10 \
#   --prop xAxisTitle="Utilization (%)" --prop yAxisTitle="Samples" \
#   --prop title.color=1F2937 --prop title.size=12 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=9 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=8:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop dataLabels=false \
#   --prop x=0 --prop y=38 --prop width=13 --prop height=18
# 5 · GPU Utilization — locked axis range so dashboard charts share scale
cli(f'add "{FILE}" "/5-ML Dashboard" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="GPU Utilization"'
    f' --prop series1=GPU:{GPU_CSV}'
    f' --prop binSize=5 --prop fill=8B5CF6'
    f' --prop axismin=0 --prop axismax=50 --prop majorunit=10'
    f' --prop xAxisTitle="Utilization (%)" --prop yAxisTitle="Samples"'
    f'{DASH}'
    f' --prop x=0 --prop y=38 --prop width=13 --prop height=18')

# officecli add charts-histogram.xlsx "/5-ML Dashboard" --type chart \
#   --prop chartType=histogram \
#   --prop title="Cost per Request ($ × 0.001)" \
#   --prop series1="Cost:<220 log-normal cost values>" \
#   --prop binSize=5 --prop overflowBin=120 --prop fill=EC4899 \
#   --prop dataLabels=true --prop "datalabels.numfmt=0" \
#   --prop xAxisTitle="Cost (m$)" --prop yAxisTitle="Requests" \
#   --prop title.color=1F2937 --prop title.size=12 --prop title.bold=true \
#   --prop title.font="Helvetica Neue" \
#   --prop axisTitle.color=6B7280 --prop axisTitle.size=9 \
#   --prop axisTitle.font="Helvetica Neue" \
#   --prop "axisfont=8:6B7280:Helvetica Neue" \
#   --prop gridlineColor=F0F0F0 \
#   --prop plotareafill=FFFFFF --prop "plotarea.border=E5E7EB:0.75" \
#   --prop chartareafill=F9FAFB --prop "chartarea.border=E5E7EB:0.75" \
#   --prop "axisline=9CA3AF:0.75" \
#   --prop x=14 --prop y=38 --prop width=13 --prop height=18
# 6 · Cost per Request — log-normal, overflow-fenced, data labels with numfmt
cli(f'add "{FILE}" "/5-ML Dashboard" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Cost per Request ($ × 0.001)"'
    f' --prop series1=Cost:{COST_CSV}'
    f' --prop binSize=5 --prop overflowBin=120 --prop fill=EC4899'
    f' --prop dataLabels=true --prop "datalabels.numfmt=0"'
    f' --prop xAxisTitle="Cost (m$)" --prop yAxisTitle="Requests"'
    f'{DASH}'
    f' --prop x=14 --prop y=38 --prop width=13 --prop height=18')


print(f"\nDone! Generated: {FILE}")
print("  6 sheets, 29 histograms total")
print("  Sheet 0 (0-Hero):              1 magazine-grade full-bleed hero poster")
print("  Sheet 1 (1-Binning Lab):       6 charts — every binning knob, identical styling")
print("  Sheet 2 (2-Distribution Zoo):  6 canonical real-world distribution shapes")
print("  Sheet 3 (3-Theme Gallery):     6 design themes on the SAME dataset")
print("  Sheet 4 (4-Typography):        4 font-family type specimens")
print("  Sheet 5 (5-ML Dashboard):      6-chart Production ML Model Report")
