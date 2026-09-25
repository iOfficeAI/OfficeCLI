---
name: officecli-pptx
description: "Use this skill any time a .pptx file is involved -- as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file; editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions 'deck', 'slides', 'presentation', 'pitch', or references a .pptx filename."
---

# OfficeCLI PPTX Skill

## Setup

If `officecli` is missing:

- **macOS / Linux**: `curl -fsSL https://d.officecli.ai/install.sh | bash`
- **Windows (PowerShell)**: `irm https://d.officecli.ai/install.ps1 | iex`

Verify with `officecli --version` (open a new terminal if PATH hasn't picked up). If install fails, download a binary from https://github.com/iOfficeAI/OfficeCLI/releases.

## ⚠️ Help-First Rule

**This skill teaches what good slides look like, not every command flag. When a property name, enum value, or alias is uncertain, consult help BEFORE guessing.**

```bash
officecli help pptx                         # List all pptx elements
officecli help pptx <element>               # Full element schema (e.g. shape, chart, animation, connector, zoom, group)
officecli help pptx <verb> <element>        # Verb-scoped (e.g. add shape, set slide)
officecli help pptx <element> --json        # Machine-readable schema
```

Help reflects the installed CLI version. When skill and help disagree, **help is authoritative**. Triggers to run help immediately: `UNSUPPORTED props:` warning, unknown animation preset, `connector.shape=` enum drifts, prop-vs-alias (`lineWidth` vs `line.width`, `color` vs `font.color`).

## Shell & Execution Discipline

**Shell quoting (zsh / bash).** ALWAYS quote element paths (`"/slide[1]/..."`) — zsh globs unquoted `[1]` to `no matches found`. Escaping happens at three layers — keep them separate (the CLI handles the second for you):

1. **Shell.** `$` in a value still belongs to the shell — single-quote the whole value: `--prop text='$15M'`. Double-quoted `"$15M"` gets expanded to `M`. The CLI does NOT unescape `\$` for you.
2. **CLI (`text=`).** The two-char escapes `\n` and `\t` ARE interpreted, consistently across pptx / docx / xlsx — `\n` is a line / paragraph break, `\t` is a tab. To produce a literal backslash-n in text, double it (`\\n`); this is rarely what you want.
3. **JSON (batch heredoc).** The recipes below pipe `cat <<EOF | officecli batch` **unquoted** so `$SLIDE` / `$FILE` expand inside the body. That same unquoted heredoc also expands a literal `$` in a value: `"$1.42"` silently becomes `.42` and `"$2.4B"` becomes `.4B` (`$1` / `$2` are empty shell params). **Escape currency as `\$`** — `"text":"\$1.42"` — which still lets `$SLIDE` expand. `\n` / `\t` inside the JSON work either way. A fully-quoted `<<'EOF'` protects every `$` but then `$SLIDE` won't expand, so only use it when the body has no shell variables. After writing money values, `view text` and confirm the `$` survived.

If in doubt, `view text` after writing and compare character-for-character.

**Incremental execution.** One command → check exit code → continue. A 50-command script that fails at command 3 cascades silently. After any structural op (new slide, chart, animation, connector) run `get` before stacking more.

## Requirements for Outputs

These are the deliverable standards every deck MUST meet. Violating any one = not done, regardless of content quality.

### All decks

**One idea per slide.** If a slide needs a second title to explain what it covers, split it. Dense "everything about X" slides lose the audience inside 3 seconds. Use a section divider to group related one-idea slides, not a mega-slide.

**Explicit type hierarchy — do NOT rely on theme defaults.** Theme defaults drift between masters. Set sizes explicitly on every text shape.

| Element | Minimum | Typical | Min shape height |
|---|---|---|---|
| Slide title | **≥ 36pt** bold | 36–44pt | ≥ 2cm |
| Section / subtitle | ≥ 20pt | 20–24pt | ≥ 1.2cm |
| Body text | **≥ 18pt** | 18–22pt | ≥ 1cm |
| Caption / axis label | ≥ 10pt muted | 10–12pt | ≥ 0.6cm |

Rule of thumb: **min shape height ≈ font_pt × 0.05cm**. An 18pt sublabel in a 0.8cm-tall box will overflow — `view annotated` catches this.

Title must be **≥ 2× body size** (36pt over 20pt works; 28pt over 20pt looks timid). Four legit exceptions to body ≥ 18pt: chart axis labels, legends, footer / page number, and ≤ 5-word KPI sublabels (e.g. "Active users"). Descriptive sentences must be ≥ 18pt. Left-align body; center only titles and hero numbers. If "the cards won't fit", drop cards instead of shrinking font.

**Two fonts max, one palette.** One heading font + one body font (e.g. Georgia + Calibri) — a third *display* face is fine only for big numerals or the cover title, as long as that heading+body pair stays intact. One dominant brand color (60–70% weight) + one supporting + one accent. Never mix 4+ colors in body content. **The palettes and font pairings in Design Principles are a floor, not a menu:** if the user gave brand colors/fonts or an existing template, match those first; otherwise the named sets are calibrated seeds — blend or diverge freely, as long as the result isn't *worse* than them and still clears the contrast floor.

**Every slide carries a non-text visual — one that informs.** Shape, chart, icon, gradient band that carries meaning, not decoration. A bullet-only deck is interchangeable with a Word doc. Exceptions: literal quote slides, code blocks, a single summary-table slide.

**Less is more — every element earns its place.** The visual rule above guards against bullet-walls; it is not licence to clutter. Don't pad with decorative stats, icons, or filler sections that don't inform ("data slop"). If a slide feels empty, fix it with layout and whitespace, not invented content — cut scope rather than bulk it up, and flag a larger addition instead of making it unprompted.

**Speaker notes on every content slide.** `--type notes --prop text="..."`. The speaker needs a script; the audience shouldn't read the slide verbatim.

**Copy reads human, not AI.** Titles orient on content, not punchline. No "It's not X. It's Y.", no manufactured tension, no faux-insight ("The magic moment"), no one-word drama ("Momentum."). Cut hype adjectives (seamless, robust, game-changing) — let the number carry it.

**Preserve existing templates.** When a file already has a theme and masters, match them. Existing conventions override these guidelines.

### Visual delivery floor (applies to EVERY deck)

Before declaring done, the per-slide render (see QA) MUST satisfy:

- **No placeholder tokens rendered as content.** `{{name}}`, `$fy$24`, `<TODO>`, `lorem`, `xxxx`, empty `()`/`[]` in chart titles never appear.
- **No overflow off-edge, no clipped text in shapes.** `view issues` flags both (`shape_off_slide` + a text-fit hint). To fix a clip: grow the box or shorten the value — never trim content to fit.
- **Cover carries its orienting elements.** Title + subtitle + presenter/client + date + a brand band or key-takeaway strap — a title-only cover reads as a stub. Generous whitespace around them is still right; rich ≠ crowded.
- **Contrast.** `view issues` auto-flags the common case — opaque dark text on a shape's own dark fill (`low_contrast`). It can't see the rest: icon / chart-series fills, scheme/inherited colors, or text over a *separate* background shape. So on any fill with brightness < 30% (`1E2761`, `36454F`, deep forest / berry / cherry), still confirm every body run, card body, chart series, and icon is `FFFFFF` or brightness > 80% — mid-gray (`6B7B8D` ≈ 44%) reads on a laptop and vanishes on projection. Spot-check via `view html` after the dark-fill pass.

If any fails, STOP and fix before declaring done.

## Design Principles

A deck is not a document. The audience has 3 seconds to get each slide. Before adding anything, ask: "If the audience reads only the biggest element and glances once, do they get the point?" If they have to read the bullets, the biggest element is wrong.

### Grid, margins, negative space

Standard widescreen is **33.87 × 19.05cm**. Treat it as a 12-column grid internally:

- **Edge margin ≥ 1.27cm** (0.5") on all sides.
- **Inter-block gap ≥ 0.76cm** (0.3") between cards / columns / rows — pick one value (0.76 or 1.27cm) and use it everywhere; mixed gaps look unfinished.
- **≥ 20% negative space per slide.** Filling every pixel reads as amateur.
- **Compose, don't web-center.** Whitespace is structural: a slide top-weighted with open space in the lower third is correct composition, not an empty defect. Intentional asymmetry (content left, breathing room right) reads more designed than centering everything — don't fill a gap just because it's there.
- For card grids: `usable = 33.87 − 2·margin − (N−1)·gap`, then `col_width = usable / N`. Don't hand-pick x coordinates.

### Font pairings

Pair by document register, not by novelty. "Best For" is a prompt, not a decree; a pairing outside this table is fine if it fits — these 8 are seeds, not the set.

| Header | Body | Best For |
|---|---|---|
| Georgia | Calibri | Formal business, finance, executive reports |
| Arial Black | Arial | Bold marketing, product launches |
| Calibri | Calibri Light | Clean corporate, minimal design |
| Cambria | Calibri | Traditional professional, legal, academic |
| Trebuchet MS | Calibri | Friendly tech, startups, SaaS |
| Impact | Arial | Bold headlines, event decks, keynotes |
| Palatino | Garamond | Elegant editorial, luxury, nonprofit |
| Consolas | Calibri | Developer tools, technical / engineering |

Set both fonts explicitly on every shape (`--prop font=Georgia` on titles, `--prop font=Calibri` on body), not via theme inheritance.

### Color and contrast

The columns: **Primary** (dominant — 60–70% of weight, the color you see first), **Secondary** (supporting tone), **Accent** (sparing, one-hit emphasis), **Text** (body on light fills), **Muted** (captions / axis labels / footer).

| Theme | Primary | Secondary | Accent | Text | Muted |
|---|---|---|---|---|---|
| Coral Energy | `F96167` | `F9E795` | `2F3C7E` | `333333` | `8B7E6A` |
| Midnight Executive | `1E2761` | `CADCFC` | `FFFFFF` | `333333` | `8899BB` |
| Forest & Moss | `2C5F2D` | `97BC62` | `F5F5F5` | `2D2D2D` | `6B8E6B` |
| Charcoal Minimal | `36454F` | `F2F2F2` | `212121` | `333333` | `7A8A94` |
| Warm Terracotta | `B85042` | `E7E8D1` | `A7BEAE` | `3D2B2B` | `8C7B75` |
| Berry & Cream | `6D2E46` | `A26769` | `ECE2D0` | `3D2233` | `8C6B7A` |
| Ocean Gradient | `065A82` | `1C7293` | `21295C` | `2B3A4E` | `6B8FAA` |
| Teal Trust | `028090` | `00A896` | `02C39A` | `2D3B3B` | `5E8C8C` |
| Sage Calm | `84B59F` | `69A297` | `50808E` | `2D3D35` | `7A9488` |
| Cherry Bold | `990011` | `FCF6F5` | `2F3C7E` | `333333` | `8B6B6B` |

Pick by topic, not by default — finance reads Midnight Executive, a product launch reads Coral Energy, safety / LOTO reads Cherry Bold. If the closest named theme is not quite right, blend (e.g. Forest primary + gold `D4A843` accent). Use **Text** on light fills, **Muted** for captions / axis / footer, `FFFFFF` or Secondary for body on dark fills.

On dark backgrounds, text and chart series follow the Hard rules contrast floor above.

### Chart-choice decision table

Wrong chart type kills the 3-second test:

| Data shape | Use | Avoid |
|---|---|---|
| Category comparison (A vs B vs C) | `column` (vertical) / `bar` (≥ 6 categories, horizontal) | pie (slices merge), line (no time axis) |
| Time series, 1–3 series | `line` | area (occlusion), bar (implies discrete) |
| Part-of-whole, 2–5 slices | `pie` / `doughnut` | pie with 8+ slices (unreadable) |
| Correlation / distribution | `scatter` | line (implies ordering) |
| Multiple categories × metrics, dense | stacked `column` or heatmap | one chart per metric — consolidate |
| KPI snapshot (single big number) | **Large-text shape** (60–72pt + ≤ 5-word sublabel), NOT a chart | gauge chart, tiny bar |

Rule of thumb: if > 3 series and > 8 categories, split into two charts or switch to a table.

### Animation

Use as much or as little as the brand and content call for — a formal finance deck trends to near-zero, a product launch can be more expressive. Animation is a tool, not décor. Three floors keep it from hurting the deck (none caps how much you use):

- **Purposeful** — each one reveals or emphasizes (progressive bullet reveal, a build-up chart), never decorates. If it doesn't aid comprehension, cut it.
- **Degrades gracefully** — pptx animation renders inconsistently across viewers (Keynote / Slides / web / mobile) and may not play at all, so every slide must read correctly as a *static* frame. Never hide essential content behind a reveal.
- **Verify live** — animation is runtime-only; `view html` and screenshots can't see it, so confirm in a real presentation viewer before shipping.

Taste steer (not a ban): `fade` / `appear` / a single `zoom-entrance` with snappy durations (~hundreds of ms) fit most decks; `bounce` / `swivel` / `spin` / `fly-from-edge` / dense multi-object choreography usually read amateur — reach for them only when the brand is deliberately playful.

### Layout patterns & data display

Vary layout across slides — repeating the same pattern makes every slide feel identical. These are common building blocks, not the full set — pick one per slide, or build a layout outside the table when the content calls for it:

| Pattern | When to use | Key measurement |
|---|---|---|
| **Two-column** (text left, visual right) | Concept + evidence; feature + screenshot | Each col ≈ 14-15cm; gap 1cm |
| **Icon rows** (icon in filled circle + bold header + description) | Feature lists, benefits, team roles | Icon circle 1.5-2cm; 3-4 rows max |
| **2×2 or 2×3 grid** (card tiles) | Quadrant analysis, SWOT, option comparison | Gap ≥ 0.76cm; consistent card height |
| **Half-bleed image** (full left or right half, content overlay on other side) | Hero moments, case study openers | Image 16-17cm wide; content column ≥ 14cm |
| **Large stat callout** (60-72pt number + ≤5-word sublabel below) | Single KPI, milestone, market size | Use shape, NOT a chart; sublabel 14-16pt muted |

**Data display quick rules:**
- Comparison columns (before/after, A vs B) beat a table for 2-3 options.
- Timelines and process flows: numbered step shapes + connectors, not a bullet list.

### Image treatment (only when a slide uses a photo / screenshot / logo)

**Read the image first** (open the file) and choose treatment from what you see — don't place blind from a filename.

- **Full-bleed photo** → size to COVER the region (crop edges), no border.
- **Screenshot / diagram / logo** → size to FIT (never crop content). A transparent or fit image sits on a contrasting fill — drop a colored rectangle behind it, don't let it float on white.
- **Text over a photo** → never raw on the image. Put it on a card, or lay a protection scrim between image and text (a dark rectangle at ~50–60% opacity, or a gradient fading from the text edge).
- Never stretch (distort the aspect ratio); don't overlay text on a busy screenshot.
- Prefer user-provided images / brand assets; no emoji or self-drawn art unless asked.

### Visual motif commitment

Pick ONE distinctive element (rounded image frames, section numbers in filled circles, single-side border band, diagonal accent strips) and carry it to every slide — commit across the whole deck; styling one slide and leaving the rest plain reads as abandoned. A secondary motif is fine only if it doesn't compete with the primary. Declare it in your build plan first: `## Motif: numbered circles in brand color`.

### Visual AI-tells to avoid

- **No decorative underline under slide titles.** A stripe / rule below a heading is the single most common AI-slide tell — use whitespace or a background-color change instead.
- **No rounded-corner card with a colored left-border accent stripe.** The other classic AI-slide tell — use a solid fill, a top accent band, or whitespace separation instead.
- **No emoji as iconography** unless the brand uses them — use a shape or a real icon asset.

Copy-level tells live in "Copy reads human".

## Common Workflow

1. **Open/save lifecycle.** `officecli open <file>` at the start, `officecli save <file>` at the end to flush your edits to disk. `save` only writes — it leaves the resident warm for any follow-up edit; reach for `officecli close <file>` only when you want to release the resident immediately (a one-shot handoff). Both are always safe — they never error or lose work. Use `batch` for repetitive shape grids. **Flush only at the non-officecli boundary:** officecli's own reads always see your edits; run `save`/`close` only before a non-officecli program reads the file (python-pptx, PowerPoint, a renderer, delivery).
2. **Orient.** New deck: `officecli create "$FILE"`. Existing: `officecli view "$FILE" outline` first. Never edit blind.
3. **Title sequence first (plan, don't build yet).** Before creating any slide or shape, write out the full ordered list of slide titles. If someone reading ONLY the titles can't follow the argument, fix the arc now — cheaper in a list than after 14 slides. Pick ONE title grammar — all topic noun-phrases or all action statements, never a mix — and hold it throughout (see "Copy reads human").
4. **Build in display order.** Add slides in audience-view order: cover → agenda → section-1 divider → section-1 content → section-2 divider → … → closing. `--index` on slide add works, but linear append keeps the build script readable and avoids index-arithmetic bugs. **Before final delivery, confirm slide count + narrative arc match your build plan.** Gate 3's order-sanity check catches cases where the cover ends up as slide 11 of 14 instead of slide 1.
5. **Incremental per slide.** Create slide + background, then title, then supporting shapes / charts / connectors. Always `layout=blank` for custom designs. After each structural op, `get /slide[N] --depth 1` to confirm shape IDs.
6. **Format to spec.** Per the Requirements table; formatting is deliverable, not polish.
7. **Save + verify.** `officecli save` flushes the file to disk (or `officecli close` to flush and also end the session). Always open in the target presentation viewer before shipping — chart colors, animations, fonts, and zoom are runtime features `view html` can't render. Full verification in QA below.
8. **QA — assume there are problems.** Fix-and-verify until a cycle finds zero new issues.

## Quick Start

Minimal viable deck: cover + one content slide + notes. `$FILE` stands in for your filename.

```bash
FILE="deck.pptx"
officecli create "$FILE"
officecli open "$FILE"

# Cover — dark fill, centered title
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761
officecli add "$FILE" /slide[1] --type shape --prop text="FY26 Strategic Review" \
  --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=3cm \
  --prop font=Georgia --prop size=44 --prop bold=true --prop color=FFFFFF --prop align=center

# Content — white fill, title + body + notes
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" /slide[2] --type shape --prop text="Revenue grew 18% YoY" \
  --prop x=1.5cm --prop y=1.2cm --prop width=30cm --prop height=2cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761
officecli add "$FILE" /slide[2] --type shape --prop text="Enterprise renewals + new EMEA region drove the beat; NRR held at 118%." \
  --prop x=1.5cm --prop y=4cm --prop width=30cm --prop height=3cm \
  --prop font=Calibri --prop size=20 --prop color=333333
officecli add "$FILE" /slide[2] --type notes --prop text="Lead with the 18% beat, preview EMEA."

officecli save "$FILE"
officecli validate "$FILE"
```

Shape of every build: open → slide+background → title → body → notes → save → validate.

## Reading & Analysis

Start wide, then narrow. `outline` first, `view text` / `get` / `query` once you know where to look.

```bash
officecli view "$FILE" outline          # slide count + titles
officecli view "$FILE" annotated        # complete per-slide breakdown with fonts, sizes, tables, charts
officecli view "$FILE" text --start 1 --end 5   # text dump (includes table cell text)
officecli view "$FILE" issues           # empty slides, overflow hints
officecli view "$FILE" stats            # counts + totals (incl. pictures missing alt)
```

**Inspect one element.** XPath-style paths, 1-based. ALWAYS quote. Prefer `@name=` / `@id=` selectors over positional `[N]` (stable across reorderings). `[last()]` works. Add `--json` for machine output.

```bash
officecli get "$FILE" "/slide[1]" --depth 1              # shape list with IDs and names
officecli get "$FILE" "/slide[1]/shape[@name=Title]"
officecli get "$FILE" "/slide[1]/table[1]" --depth 3     # table rows / cells
```

**Query across the deck.** CSS-like selectors; operators `=`, `!=`, `~=`, `>=`, `<=`, `[attr]`, `:contains()`, `:no-alt`. `help pptx query` lists queryable element types.

```bash
officecli query "$FILE" 'shape:contains("Revenue")'
officecli query "$FILE" 'picture:no-alt'                 # accessibility gap
officecli query "$FILE" 'shape[fill=1E2761]'             # color match
officecli query "$FILE" 'shape[width>=10cm]'             # numeric
```

**`query --json` output schema.** Results wrap in `.data.results[]` — `jq -r '.data.results[0].format.id'`, NOT `.[0].id`. Shape name is `.name`; fill is `.format.fill`; textColor is `.format.textColor`.

**Visual preview (LEAD).**

```bash
officecli view "$FILE" html                # prints an HTML preview path; Read it for per-slide visual audit (best structural ground truth)
officecli view "$FILE" svg --start 3 --end 3   # single slide SVG (charts + gradients do NOT render in SVG)
```

**Reading the output — an expected non-defect:**
- **`layout=blank` has no title placeholder.** Titles are plain `shape` elements, so `view outline` reporting `(untitled)` is **expected**, not a defect. Use `layout=title` + `placeholder[title]` only when screen-reader outline compatibility matters.

## Creating & Editing

Verbs: `add` / `set` / `remove` / `move` / `swap` / `batch` / `raw-set`. Ninety percent of a deck is slides, shapes, text, a few charts, pictures, connectors.

### Slides and backgrounds

A slide is `/slide[N]`. Always pass `layout=blank` for custom designs. Background: solid, gradient, or image.

```bash
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761                 # solid
officecli add "$FILE" / --type slide --prop layout=blank --prop "background=1E2761-CADCFC-180"   # gradient (start-end-angle)
officecli add "$FILE" / --type slide --prop layout=blank --prop background=image:/path/to/hero.jpg  # image background (LEAD)
```

### Shapes

A `shape` holds text, fill, border, position, and optional animation / link.

```bash
officecli add "$FILE" /slide[2] --type shape --prop name=Title --prop text="Key Insight" \
  --prop x=2cm --prop y=2cm --prop width=20cm --prop height=3cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
```

Positioning is explicit — no layout engine, you own the grid math. `--prop preset=` picks geometry (`rect`, `roundRect`, `ellipse`, `triangle`, `arrow`, `star5`, ...); custom `M...Z` paths are not supported — pick a preset. **Name shapes at creation** (`--prop name=HeroTitle`) and address later with `"/slide[N]/shape[@name=HeroTitle]"` — names survive z-order / remove-then-add, whereas positional `/shape[3]` (and even `@id=`) shift. Re-`get --depth 1` after any structural change before using positional indexes.

### Geometric align & distribute (one call — never hand-compute)

Creation still needs grid math, but **fixing排版 never does**. The `layout` verb repositions shapes server-side in a single call — no get-coordinates → arithmetic → N×set loop:

```bash
# Bottom-align every auto-shape on slide 2, whatever their current y/height is:
officecli layout "$FILE" /slide[2] --align bottom
# Equal horizontal gaps across 3+ shapes (endpoints stay fixed):
officecli layout "$FILE" /slide[2] --distribute horizontal
# Both at once + only some shapes (stable @id paths preferred):
officecli layout "$FILE" /slide[2] --align bottom --distribute horizontal \
  --targets "shape[@id=100000],shape[@id=100001]"
```

Rules: (1) `layout` takes a SLIDE path (`/slide[N]`); shape paths are rejected. Do NOT invent `align`/`distribute` as verbs and do NOT set geometric `align=` on a shape path — that is TEXT alignment (left|center|right|justify), a different operation. (2) `--align` values: `left|center|right|top|middle|bottom` (relative to the targeted shapes' bounding box) or `slide-left|slide-center|...` (relative to the slide). (3) `--distribute horizontal|vertical` needs ≥3 shapes. (4) `--targets` accepts `shape[N]`, bare `N`, or `shape[@id=N]`; omit = all auto-shapes/textboxes. (5) LIMIT: pictures/charts/tables/connectors are NOT matched — only `p:sp` shapes; place those by hand. (6) Equivalent legacy form (same engine): `set "$FILE" /slide[2] --prop align=bottom` — prefer `layout`. (7) Always verify with `get --depth 1` (x/y readback; values may come back as `Ncm` or `Nemu`) or `view screenshot` after. (8) Delivery PDF: `view "$FILE" pdf -o out.pdf` — headless-Chrome fallback, no exporter plugin needed. Works mid-session (it releases a warm resident itself); `close` first is still good hygiene before handing off.

### Text inside shapes (paragraphs, runs, styling)

A shape has paragraphs (`paragraph[K]`) and runs (`run[K]`). For one-line text, `--prop text=` on the shape is enough; a `\n` in the text makes a paragraph break, `\t` a tab (see Shell & Execution Discipline; double `\\n` for a literal). `add --type paragraph` takes the same style props as a shape (text, align, bold, italic, size, color, font). For mixed styling *within* a line, append a styled run:

```bash
officecli add "$FILE" "/slide[2]/shape[@name=Card1]/paragraph[1]" --type run \
  --prop text=" (inline detail)" --prop size=14 --prop italic=true --prop color=8899BB
```

### Charts

Pick chart type per the Design Principles chart-choice table. Full prop list (chartType enum, `seriesN.*`, `data=`/`categories=`, axis options): `help pptx add chart`. Typical multi-series with brand colors:

```bash
officecli add "$FILE" /slide[3] --type chart --prop chartType=column \
  --prop series1.name=Revenue --prop series1.values="42,45,48" --prop series1.color=1E2761 \
  --prop series2.name=Growth  --prop series2.values="2,7,7"    --prop series2.color=CADCFC \
  --prop categories="Q1,Q2,Q3" \
  --prop x=2cm --prop y=4cm --prop width=20cm --prop height=10cm
```

Gotchas: (1) chart titles with `()`, `[]`, `TBD` ship as literal text. (2) some viewers normalize chart colors to theme defaults — verify in the target viewer. Series can be added after creation (`add --type series`).

### Pictures

```bash
officecli add "$FILE" /slide[4] --type picture --prop src=hero.jpg \
  --prop x=1cm --prop y=1cm --prop width=32cm --prop height=18cm \
  --prop alt="Product hero, gradient lit from right"
```

Confirm with `officecli query "$FILE" 'picture:no-alt'` — must be empty before delivery.

### Connectors (LEAD — flowcharts / decision trees first-class)

Draws a line between two shapes or free coordinates. Full prop / enum reference (`shape`, `headEnd`/`tailEnd` values, `from`/`to` ref forms): `help pptx add connector`.

```bash
officecli add "$FILE" /slide[5] --type connector \
  --prop "from=/slide[5]/shape[@name=BoxA]" --prop "to=/slide[5]/shape[@name=BoxB]" \
  --prop shape=elbow --prop color=333333 --prop tailEnd=triangle
```

**Every flow connector needs an arrowhead.** Without one, `bentConnector3` renders as a directionless line. `preset=rightArrow` overlay only works for horizontal flows; diamonds / decision trees with diverging edges need `tailEnd=`.

### Animations (LEAD)

Use per the Animation floors above (purposeful, degrades gracefully, verify live). Preset names + duration syntax: `help pptx animation`.

```bash
officecli set "$FILE" "/slide[2]/shape[@name=HeroCard]" --prop animation=fade-entrance-400
officecli set "$FILE" "/slide[2]/shape[@name=HeroCard]" --prop animation=none    # clear all
```

### Hyperlinks, tooltips, slide-jump

`--prop link=slide[N]` for an in-deck jump (1-based; target slide must exist), `link=nextslide` / `firstslide` / `lastslide` / `previousslide` / `endshow` for named navigation, `link=https://...` for a URL, `--prop tooltip="..."` for hover text.

### Tables, placeholders, groups, zoom — one-liners

- **Tables** — `--type table --prop rows=N --prop cols=M`. Row-level `set` supports `height` and `c1/c2/c3` (seed cell text). Header-row styling is table-level (`firstRow=true` / `headerFill=`), not a row prop. Cell formatting lives on the cell paragraph / run. Populate rows BEFORE setting table-level font (font cascade gets reset by row ops).
- **Placeholders** — `"/slide[N]/placeholder[title]"` / `placeholder[body]`. Available only when the slide uses a layout with placeholders (not `layout=blank`).
- **Groups** (LEAD) — address children via `"/slide[N]/group[@name=G]/shape[1]"`. Survives reordering better than positional indexes.
- **Zoom slide** (LEAD) — `--type zoom --prop target=N` (one link per target; alias `slide`). Emit N separate zoom shapes for a multi-target nav hub. Zoom is a runtime feature — `view html` shows the static geometry; the zoom interaction runs only in a live presentation viewer.
- **Slide comments** — reviewer annotations anchored at `/slide[N]/comment[M]`. Full lifecycle (`add / set / get / query / remove`). Props: `text`, `author`, `initials` (auto-derived), `date` (ISO 8601, defaults to UtcNow), `x` / `y` (length anchor).
  ```bash
  officecli add "$FILE" "/slide[2]" --type comment --prop author="Alice" --prop text="Tighten this bullet" --prop x=20cm --prop y=3cm
  officecli query "$FILE" 'comment' --json | jq '.data.results | length'   # count all review comments
  officecli remove "$FILE" "/slide[2]/comment[1]"                           # resolve after addressing
  ```

### Deck-level recipes

Patterns not obvious from the primitives. Each gives the **visual outcome** first, then a runnable block. `$FILE` = your filename. Use `/slide[last()]` to address the slide you just added. The recipes demonstrate **structure and coordinate math** — swap in the palette / fonts you chose for this topic; the navy `1E2761` + Georgia is just the example's theme, not a house style to copy verbatim.

**Z-order.** Later-added shapes are on top. Add background decoration FIRST, titles LAST. To fix after the fact: `--prop zorder=back/front` (renumbers siblings — re-`get --depth 1` before stacking more).

#### (a) Cover (and section divider)

**Visual outcome.** Dark navy fill, centered 44pt title, 18pt ice-blue meta line.

```bash
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761
officecli add "$FILE" "/slide[last()]" --type shape --prop text="Strategic Growth Review" \
  --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=3cm \
  --prop font=Georgia --prop size=44 --prop bold=true --prop color=FFFFFF --prop align=center
officecli add "$FILE" "/slide[last()]" --type shape --prop text="Prepared for Acme Leadership — FY26 Outlook" \
  --prop x=2cm --prop y=11cm --prop width=29.87cm --prop height=1.2cm \
  --prop font=Calibri --prop size=18 --prop color=CADCFC --prop align=center
```

**Section divider** = same cover, plus a giant translucent number (`size=120`, `opacity=0.15`) added FIRST so it sits behind the section title.

#### (b) Data slide (chart + commentary block)

**Visual outcome.** Left two-thirds: column chart with brand series colors. Right one-third: "Key Insight" card with 20pt heading + 18pt body — audience reads the takeaway before parsing the bars.

```bash
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[last()]" --type shape --prop text="FY26 Revenue Beat Plan by 18%" \
  --prop x=1.5cm --prop y=1cm --prop width=30cm --prop height=1.8cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761

# Chart — left 2/3 (single-quote the title because of `$`)
officecli add "$FILE" "/slide[last()]" --type chart --prop chartType=column \
  --prop series1.name=Actual --prop series1.values="42,45,48,55" --prop series1.color=1E2761 \
  --prop series2.name=Plan --prop series2.values="40,42,45,48" --prop series2.color=CADCFC \
  --prop categories="Q1,Q2,Q3,Q4" --prop x=1.5cm --prop y=3.5cm --prop width=20cm --prop height=14cm --prop title='FY26 Revenue ($M)'

# Commentary card — right 1/3: background + heading + body
officecli add "$FILE" "/slide[last()]" --type shape --prop preset=roundRect --prop fill=F5F7FA --prop line=none \
  --prop x=22.5cm --prop y=3.5cm --prop width=9.8cm --prop height=14cm
officecli add "$FILE" "/slide[last()]" --type shape --prop text="Key Insight" \
  --prop x=23cm --prop y=4cm --prop width=9cm --prop height=1.2cm \
  --prop font=Georgia --prop size=20 --prop bold=true --prop color=1E2761
officecli add "$FILE" "/slide[last()]" --type shape --prop text="EMEA launch + NRR at 118% drove 12pp of the 18pp beat." \
  --prop x=23cm --prop y=5.5cm --prop width=9cm --prop height=11cm \
  --prop font=Calibri --prop size=18 --prop color=333333
```

#### (c) Flowchart / process diagram (boxes + connectors)

**Visual outcome.** Four rounded boxes across at y=8cm, each 6×3cm, alternating navy/iceblue, joined by elbow connectors with triangle arrowheads.

Grid math (4 boxes, 33.87cm slide, 1.5cm margins): `gap = (33.87 − 3 − 24) / 3 = 2.29cm`. x-positions: `1.5, 9.79, 18.08, 26.37`.

Each box carries its own label via `valign=middle` (no separate overlay shape needed). Use `batch` heredoc for portable coordinate arithmetic — no `bc`, no bash arrays.

```bash
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step1","preset":"roundRect","fill":"1E2761","line":"none","x":"1.5cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 1","font":"Georgia","size":"20","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step2","preset":"roundRect","fill":"CADCFC","line":"none","x":"9.79cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 2","font":"Georgia","size":"20","bold":"true","color":"1E2761","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step3","preset":"roundRect","fill":"1E2761","line":"none","x":"18.08cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 3","font":"Georgia","size":"20","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step4","preset":"roundRect","fill":"CADCFC","line":"none","x":"26.37cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 4","font":"Georgia","size":"20","bold":"true","color":"1E2761","align":"center","valign":"middle"}}
]
EOF

# Connector pattern — reuse for any box-to-box graph.
for pair in "Step1 Step2" "Step2 Step3" "Step3 Step4"; do
  A=${pair% *}; B=${pair#* }
  officecli add "$FILE" "/slide[$SLIDE]" --type connector \
    --prop "from=/slide[$SLIDE]/shape[@name=$A]" \
    --prop "to=/slide[$SLIDE]/shape[@name=$B]" \
    --prop shape=elbow --prop color=333333 --prop tailEnd=triangle
done
```

`shape=elbow` is canonical (`bentConnector2` / `bentConnector3` also accepted).

#### (d) Multi-slide deck skeletons

No code block — it's a rhythm. The sequences below are **illustrations of one working cadence (alternating dark dividers with white content), not required running orders** — derive your actual arc from the content first (see "Title sequence first"), then borrow whatever divider/content rhythm fits:

- **10-slide review:** Cover · Agenda · 3 KPI · Div01 · Chart · Chart · Div02 · Flow · Timeline · Close
- **20-slide pitch:** same rhythm × 2, sectioned Problem · Solution · Market · Product · Traction · Model · Team · Financials · Ask
- Every divider must appear **before** its section content (Gate 3 order sanity)
- Cover/divider = (a); chart pages = (b); process pages = (c); KPI pages = (e); decision pages = (f)

#### (e) KPI callouts — giant-number card grid

**Visual outcome.** Three or four giant numbers across a row; each card = unit sublabel + small percent-change chip + one-line takeaway. The single most common exec-deck element.

**Sizing rule.** 60pt Georgia bold fits ~5 chars in a 9.78cm card (`$84.2`, `118%`, `24.5`). For longer values (`$84.2M`), split: `$84.2` as the big number, `USD millions` as the sublabel — never shrink the font to chase a unit suffix, it just wraps.

Grid math (3 cards, 1.5cm margins, 0.76cm gap): `col_width = (33.87 − 3 − 1.52) / 3 = 9.78cm`. x-positions: `1.5, 12.04, 22.58`. Use accent color on a single "watch" card so risk reads in one second.

```bash
# Two cards: navy standard + terracotta watch. Each = bg + big number + sublabel + chip.
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"preset":"roundRect","fill":"1E2761","line":"none","x":"1.5cm","y":"4cm","width":"9.78cm","height":"7cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"84.2","x":"1.5cm","y":"4.8cm","width":"9.78cm","height":"2.8cm","font":"Georgia","size":"60","bold":"true","color":"FFFFFF","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"USD millions · ARR","x":"1.5cm","y":"8cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","color":"CADCFC","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"+24% YoY","x":"1.5cm","y":"9cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","bold":"true","color":"CADCFC","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"preset":"roundRect","fill":"B85042","line":"none","x":"22.58cm","y":"4cm","width":"9.78cm","height":"7cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"\$1.42","x":"22.58cm","y":"4.8cm","width":"9.78cm","height":"2.8cm","font":"Georgia","size":"60","bold":"true","color":"FFFFFF","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"CAC payback (yrs)","x":"22.58cm","y":"8cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","color":"FFFFFF","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"+8% — watch","x":"22.58cm","y":"9cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","bold":"true","color":"FFFFFF","align":"center"}}
]
EOF
```

#### (f) Decision tree — YES/NO branching

**Visual outcome.** Diamond at top-center; YES/NO child boxes diverging left-right; both converge into a shared terminal box. Layout: diamond at `x=13.94, y=2cm, 6×3cm`; YES at `3cm, 7.5cm`; NO at `22.87cm, 7.5cm`; terminal at `13.94cm, 13cm`. Convention: red = stop/escalate, blue = standard, green = safe terminal. **Every connector needs an arrowhead** — readers misparse direction otherwise.

```bash
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Decide","preset":"diamond","fill":"1E2761","line":"none","x":"13.94cm","y":"2cm","width":"6cm","height":"3cm","text":"Hazardous energy present?","font":"Calibri","size":"14","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"YesBox","preset":"roundRect","fill":"B85042","line":"none","x":"3cm","y":"7.5cm","width":"8cm","height":"3cm","text":"Lockout + Tagout + Verify","font":"Calibri","size":"16","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"NoBox","preset":"roundRect","fill":"CADCFC","line":"none","x":"22.87cm","y":"7.5cm","width":"8cm","height":"3cm","text":"Proceed with standard PPE","font":"Calibri","size":"16","bold":"true","color":"1E2761","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Done","preset":"roundRect","fill":"2C5F2D","line":"none","x":"13.94cm","y":"13cm","width":"6cm","height":"2.5cm","text":"Begin service","font":"Calibri","size":"16","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}}
]
EOF
```

Then 4 connectors (`Decide→YesBox`, `Decide→NoBox`, `YesBox→Done`, `NoBox→Done`) using the connector loop pattern from (c).

## QA (Required)

**Assume there are problems.** First render is almost never correct. If you found zero issues, you were not looking hard enough.

### Delivery Gate (any failure = REJECT, do NOT deliver)

Gates 1–2b are text/schema-level (cannot see a rendered slide); Gate 3 is the only visual check. Done = every gate PASS **and** Gate 3 loop converged.

Each gate is **run a command, judge its output** — the officecli commands are identical on every OS (macOS / Linux / Windows), so no shell scripting is needed; the judging is yours.

- **Gate 1 — schema.** `officecli validate "<file>"`. Any schema error → REJECT and fix.
- **Gate 2 — overflow / format / structure.** `officecli view "<file>" issues`. If it lists *any* issue (lines tagged `[O1]`, `[C1]`, `[S1]`, …) → REJECT, fix, re-run until clean.
- **Gate 2b — leftover placeholders.** `officecli view "<file>" text`, then scan the output for `xxxx`, `lorem` / `ipsum`, `<TODO>`, `placeholder`, "this slide layout", or empty `()` / `[]`. Any hit → REJECT.

### Gate 3 — Visual audit (MANDATORY)

Pick **one** path:

**Screenshot (default)** — for vision-capable agents. Screenshot each slide in turn — `officecli view "<file>" screenshot --page 1 -o slide1.png`, then `--page 2`, … — until the page index runs past the deck (one screenshot = one slide). If it errors on page 1, use the fallback below.

**Judge every PNG against the checklist, adversarially** — "assume problems exist; finding none means you didn't look hard enough." Report one `slide N: <issue>` line per problem, or `PASS`. This step is required however you run it. **If** your harness can spawn a subagent, delegate the judging to a *fresh, independent* one — the agent that built the deck is biased toward "looks fine", a separate pair of eyes is more critical — handing it the screenshots + this checklist and the same adversarial framing. No subagent? Do exactly the same yourself.

**Fallback — HTML-text** (no vision, or screenshot failed): read `view "$FILE" html` as text. DOM cannot prove **dark-on-dark / fine overlap / arrowheads / gap-margin metrics / column alignment** — flag these as "not visually verified" rather than PASS.

**Optional `--grid N`** — only on user request for layout-rhythm, or when `view outline` shows anomalous layout distribution: `officecli view "<file>" screenshot --grid 3 -o grid.png`.

**Per-slide checklist (assume issues exist):**

- **overlap** — shapes / charts / giant decorative numbers (01/02/03 100pt+) colliding
- **text overflow** — clipped at slide or shape boundary (KPI cards, narrow boxes)
- **narrow text box** — content fits technically but wraps to many short lines (1–2 words each); long sublabel in a 3cm KPI card, body line in a too-tight column
- **dark-on-dark** — fill brightness < 30% with text/icon brightness < 80% (incl. dark icons on dark without a contrasting circle)
- **image treatment** — photo stretched/distorted, text raw on a busy image (no card/scrim), screenshot or logo cropped, transparent image floating on white
- **missing arrowheads** — flowchart connectors as plain lines
- **decorative-line / title mismatch** — accent bar sized for one-line title but title wrapped to two (or vice versa)
- **footer / citation collision** — source line, page number, or footnote touching content above
- **tight margin / gap** — element within ~0.5" of slide edge, or two cards within ~0.3"
- **uneven gaps** — large empty area on one side, cramped on another (broken rhythm)
- **column / repeat-element misalignment** — KPI cards / icons off baseline or inconsistent width
- **order sanity** — sequence matches narrative (cover → agenda → dividers-before-sections → closing)

REJECT with `slide N: <issue>` lines, else "Gate 3 PASS" (HTML-text fallback adds "<unverified-items> not visually verified").

**Fix-verify (mandatory, max 3 cycles).** Fix → re-run Gate 3 → repeat until zero new issues; one fix often surfaces another. After 3 rounds without convergence, **stop** — likely seesaw, template-level cause, or agent misread. Report `slide N: <issue> — attempted: <fixes> — likely root: <template|design-conflict|ambiguous>` and let the user decide.

**Then flush (part of the gate).** Once Gate 3 converges, end with `officecli save "<file>"` — this guarantees your edits are written to disk before delivery (use `officecli close "<file>"` instead to also release the resident on a one-shot handoff). Required final step, not optional. Always safe: never errors or loses work.

## Common Pitfalls

Sanity-check cheatsheet — what breaks on the first try. Design + shell traps.

| Pitfall | Correct approach |
|---|---|
| Unquoted `[N]` in zsh/bash | Always quote paths: `"/slide[1]"`. zsh globs unquoted `[1]` → `no matches found` — #1 first-use stumble |
| `--name "foo"` | All attributes go through `--prop`: `--prop name="foo"` |
| `/shape[myname]` (bare name in brackets) | Use `@name=` selector: `/shape[@name=myname]` or `/shape[@id=10007]` |
| Paths 1-based vs `--index` 0-based | `/slide[1]` = first slide; `--index 0` = first position |
| `$` in `--prop text=` | Single-quote: `--prop text='$15M'`. Double-quoted `"$15M"` gets shell-expanded to `M` |
| `\n` / `\t` in `--prop text=` | Interpreted by the CLI: `\n` = paragraph break, `\t` = tab. Double `\\n` for a literal |
