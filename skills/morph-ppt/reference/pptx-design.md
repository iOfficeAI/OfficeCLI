---
name: pptx-design
description: Morph-specific design notes — Scene Actors + Page Types + Shape Index + Morph Animation Essentials
---

# Morph Design Essentials

Canvas / fonts / colors / contrast → see `skills/officecli-pptx/SKILL.md` §Requirements / §Design Principles / §Visual delivery floor. This file covers only the morph-specific material absorbed into SKILL.md's §Scene Actors / §Choreography / §Morph Pair Planning, plus detail that didn't fit there.

---

## 1) Scene Actors (Animation Engine) — expanded

**Purpose.** Create smooth Morph animations through persistent shapes that change properties across adjacent slides.

### Setup

Define 6-8 actors on Slide 1 if the deck tells a continuous-visual story:

- **Large** (5-8cm): Main visual anchors (hero circle, band, hero card)
- **Medium** (2-4cm): Supporting elements (metric cards, accent rings)
- **Small** (1-2cm): Accents and details (dots, dashes, icons)

**Shape types** available via `--prop preset=`: `ellipse | rect | roundRect | triangle | diamond | star5 | hexagon`. Full list: `officecli help pptx shape`.

### Naming (SKILL.md is authoritative)

Three-prefix system — `!!scene-*` / `!!actor-*` / `#sN-*`. Source of truth: `SKILL.md` §What is Morph? — core mechanics. This file adds only the Python-vs-shell quoting note below.

**Python:** `#` and `!!` require no special quoting — pass as plain strings in `subprocess.run([..., "--prop", "name=#s1-title", ...])`.

**Shell (bash/zsh):** ALWAYS single-quote to avoid history expansion on `!!` and comment-leading on `#`: `--prop 'name=!!scene-ring'` / `--prop 'name=#s1-title'`.

### Pairing example — 3 actors × 3 slides

```
Slide 1: !!scene-ring (x=5cm, y=3cm, w=8cm, fill=E94560, opacity=0.3)
         !!scene-dot  (x=28cm, y=15cm, w=1cm)
         !!actor-headline (x=4cm, y=8cm, w=26cm, size=48)

Slide 2: !!scene-ring (x=20cm, y=2cm, w=12cm, opacity=0.6)   ← same name, new position+size
         !!scene-dot  (x=3cm, y=16cm, w=1.5cm)                ← moved to opposite corner
         !!actor-headline (x=1.5cm, y=1cm, w=12cm, size=24)  ← shrunk + moved to top-left

Slide 3: !!scene-ring (x=36cm)                                ← ghosted off-canvas
         !!scene-dot  (x=10cm, y=2cm, w=1cm)
         !!actor-headline (x=36cm)                            ← ghost: new headline takes over
         !!actor-subpoint (x=4cm, y=8cm, w=26cm, size=36)    ← new actor enters (no pair on S2 = fade in)
```

### Per-slide content (`#sN-*`) workflow

1. **Clone previous slide** → inherited `#s(N-1)-*` content carries the old slide's prefix.
2. **Ghost inherited content** → move all `#s(N-1)-*` shapes to `x=36cm`.
3. **Add new content** → with current slide's prefix `#sN-*`.

Without step 2, slides accumulate shapes → visual overlap compounds silently across the deck.

---

## 2) Page Types (mix for rhythm)

Vary page types to avoid monotony. Each serves a different narrative purpose:

| Type | When to use | Visual structure |
|---|---|---|
| **hero** | Opening, closing | Large centered title + scattered scene actors |
| **statement** | Key message, transition | One impactful sentence + dramatic actor shifts (8cm+ moves) |
| **pillars** | Multi-point structure | 2-4 equal columns, actors become card backgrounds (opacity 0.12) |
| **evidence** | Data, statistics | 1-2 large asymmetric blocks + supporting details (opacity 0.3-0.6) |
| **timeline** | Process, sequence | Horizontal or vertical flow with step backgrounds |
| **comparison** | A vs B | Left-right split (50/50 or 60/40) with contrasting colors |
| **grid** | Multiple items | Scattered or grid layout, lighter feel |
| **quote** | Breathing moment | Centered text, minimal decoration |
| **cta** | Call to action | Return to bold, centered design |
| **showcase** | Featured display | Large central area for product/screenshot |

**Design notes:**

- **pillars**: Multi-column even distribution; scene actors morph into card backgrounds (roundRect, opacity=0.12).
- **evidence**: Asymmetric — 1 large actor (30-40% canvas) + 1 medium (20-30%), opacity 0.3-0.6 allowed for data backgrounds.
- **grid**: Must differ from pillars and evidence — light, scattered vs. structured.
- **Variety matters**: Avoid repeating the same page type consecutively.

---

## 3) Shape Index Mechanics

Shapes are numbered sequentially on each slide: `shape[1]`, `shape[2]`, `shape[3]`... When `transition=morph` is applied, CLI auto-prefixes `!!` to names — **use index paths after that** (see SKILL.md §Known Issues M-1).

### Index behavior

- **On creation:** Shapes added in order get increasing indices.
- **After cloning:** New slide inherits all shapes with identical indices.
- **After adding to a cloned slide:** New shapes get the next available index.
- **After modifying:** Index stays the same.

### Pattern for build scripts

```
Slide 1: 6 actors + 2 content = 8 shapes total
Slide 2: Clone (8) → Ghost content (shape[7-8]) → Add new (shape[9+])
Slide 3: Clone (10) → Ghost content (shape[9-10]) → Add new (shape[11+])
```

**Formula:** Next slide's first new shape index = Previous slide's total shape count + 1.

**Debugging:** `officecli get <file> '/slide[N]' --depth 1` to inspect actual indices.

---

## 4) Morph Animation Essentials

### Minimum requirements

1. Slides 2+ must have `transition=morph` (`officecli set /slide[N] --prop transition=morph`).
2. Scene actors must have identical `name=` across slides.
3. Previous per-slide content must be ghosted (`x=36cm`) before adding new content.
4. Adjacent slides should have different spatial layouts (displacement ≥ 5cm OR rotation ≥ 15° OR size delta ≥ 30% on ≥ 3 shapes).

### Creating motion

Change ≥ 3 scene-actor properties between adjacent slides:

- Move positions (x, y)
- Resize (width, height)
- Rotate (rotation degrees)
- Shift colors (fill, opacity)

**Goal:** Sense of movement + transformation, not just fade.

### Entrance effects on morph slides

Morph handles shape transitions automatically — entrance animations are usually unnecessary. If one is needed (e.g., fade a new `#sN-*` card in), use the `with` trigger so it plays simultaneously with morph:

```
animation=fade-entrance-300-with
```

Format: `EFFECT[-DIRECTION][-DURATION][-TRIGGER]`. See `officecli help pptx animation` for preset list.

---

## 5) Style References

52 visual style directories in `reference/styles/` — see `reference/styles/INDEX.md` for the catalog. Lookup workflow is in SKILL.md §Style library lookup workflow. Key rule: **learn the approach, do not copy coordinates** (the style build.sh files have known typesetting bugs per `INDEX.md` L5-11).
