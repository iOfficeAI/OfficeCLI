# Deck semantic controls (CSBU WorkMate)

Catalog-declared slide controls honored by OfficeCLI packing/render and PresentationStudio Inspector.
Unknown control ids are ignored (do not fail validate/build).

## Core keys (P0–P1)

| id | type | effect |
|----|------|--------|
| `moduleCount` | range | How many module slots (cards/metrics/steps/…) stay visible and row-packed |
| `balance` | range | Left/right width split for two-column layouts |
| `mediaSide` | select `left\|right` | Mirror media/text for split image covers |
| `showInsight` | toggle | Hide `insight` slot; expand companion chart/table/matrix |
| `chartType` | select | Overlay chart block `chartType` at build time |
| `emphasis` | select | Title emphasis hint (`subtle\|bold\|hero`) |

Per-slot: `slot.<id>.visible` (bool) generic visibility (P0.1).

## P2 keys (wave5 / 1.0.160+)

| id | type | effect |
|----|------|--------|
| `density` | select `compact\|comfortable\|spacious` | Row-pack start/total/gap metrics |
| `columns` | range | Preferred visible column count for `col*` packs (falls back to `moduleCount`) |
| `showLegend` | toggle | Chart prop `legend` = `right` or `none` |
| `showAxisLabels` | toggle | Chart prop `axisVisible` true/false |
| `showCallout` | toggle | Hide `callout` slot (same pattern as insight) |
| `showFooter` | toggle | Hide `footer` slot |

Do not invent unused keys to pad counts — only ship controls that change packing or OOXML output.
