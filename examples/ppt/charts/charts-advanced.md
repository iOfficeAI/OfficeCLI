# Advanced Charts Showcase — long-tail properties

Three files work together:

- **charts-advanced.py** — build script (`python3 charts-advanced.py`).
- **charts-advanced.pptx** — generated 8-slide deck.
- **charts-advanced.md** — this file.

Covers properties **not** demonstrated by the per-type decks
(`charts-column`, `charts-bar`, …). After running every script in this
family, the only chart / chart-axis / chart-series properties that remain
unstamped are pure **get-only readback** fields — they cannot be Set as
input, but they surface in the JSON readback shapes on slide 8.

## Slide map

```
Slide 1  RTL & anchor          direction=rtl (Set), anchor named-token, anchor cm-form
Slide 2  Axis shortcuts        axisvisible / valaxisvisible / catAxisVisible,
                               axisorientation, axisposition,
                               cataxisline / valaxisline
Slide 3  Crossings             crossBetween (between/midCat), crosses (autoZero/max/min), crossesAt
Slide 4  Category axis layout  labeloffset (100/300), ticklabelskip (2/3)
Slide 5  Marker size & fills   markersize standalone, areafill, chartFill, plotvisonly
Slide 6  Style + blanks        style=2/42, dispBlanksAs=gap (Set), dataRange syntax, catTitle
Slide 7  chart-axis Set        dispUnits, logBase, minorUnit, visible, labelRotation (per-axis)
Slide 8  Get readback          chart-series get --json (alpha, outlineColor, scatterStyle, ...)
                               chart-axis   get --json (axisFont, axisMax, axisMin, axisNumFmt,
                                                        axisOrientation, axisTitle, labelOffset,
                                                        tickLabelSkip)
```

## Get-only readback fields (no Set, surface in `get --json`)

| Element | Get-only properties |
|---|---|
| chart | `id` |
| chart-axis | `axisFont`, `axisMax`, `axisMin`, `axisNumFmt`, `axisOrientation`, `axisTitle`, `labelOffset`, `tickLabelSkip` |
| chart-series | `alpha`, `categoriesRef`, `dataLabels.numFmt`, `dataLabels.separator`, `errBars`, `invertIfNeg`, `nameRef`, `outlineColor`, `scatterStyle`, `secondaryAxis` |

Slide 8 calls `officecli get --json` on a series and an axis after a known
Set, then stamps the JSON onto the slide — the readback fields appear in
that block.

## Regenerate

```bash
cd examples/ppt
python3 charts-advanced.py
# → charts-advanced.pptx
```
