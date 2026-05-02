---
name: officecli-pptx
description: "Use this skill any time a .pptx file is involved -- as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file; editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions 'deck', 'slides', 'presentation', 'pitch', or references a .pptx filename."
---

# OfficeCLI PPTX Skill

## BEFORE YOU START (CRITICAL)

**If `officecli` is not installed:**

`macOS / Linux`

```bash
if ! command -v officecli >/dev/null 2>&1; then
    curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
fi
```

`Windows (PowerShell)`

```powershell
if (-not (Get-Command officecli -ErrorAction SilentlyContinue)) {
    irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
}
```

Verify: `officecli --version`

If `officecli` is still not found after first install, open a new terminal and run the verify command again.

If the install command above fails (e.g. blocked by security policy, no network access, or insufficient permissions), install manually — download the binary for your platform from https://github.com/iOfficeAI/OfficeCLI/releases — then re-run the verify command.

## ⚠️ Help-First Rule

**This skill teaches what good slides look like, not every command flag. When a property name, enum value, or alias is uncertain, consult help BEFORE guessing.**

```bash
officecli help pptx                         # List all pptx elements
officecli help pptx <element>               # Full element schema (e.g. shape, chart, animation, connector, zoom, group, background)
officecli help pptx <verb> <element>        # Verb-scoped (e.g. add shape, set slide)
officecli help pptx <element> --json        # Machine-readable schema
```

Help reflects the installed CLI version. When skill and help disagree, **help is authoritative**. Triggers to run help immediately: `UNSUPPORTED props:` warning, unknown animation preset, `connector.shape=` enum drifts, prop-vs-alias (`lineWidth` vs `line.width`, `color` vs `font.color`).

## Shell & Execution Discipline

**Shell quoting (zsh / bash).** ALWAYS quote element paths (`"/slide[1]/..."`) — zsh globs unquoted `[1]` to `no matches found`. CLI does NOT interpret `\$` / `\t` / `\n`; escapes happen at three layers, only one of which is yours:

1. **Shell.** `$` in a value → single-quote the whole value: `--prop text='$15M'`. Double-quoted `"$15M"` gets shell-expanded to `M`.
2. **JSON (batch).** Real newline in shape text goes via `"\n"` inside a `<<'EOF'` heredoc. Writing `\n` in shell-quoted `--prop text=` is a bug.
3. **pptx.** A line break is `<a:br/>`, not `\n`. Prefer multiple `--type paragraph` adds over fighting escapes.

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

Title must be **≥ 2× body size** (36pt over 20pt works; 28pt over 20pt looks timid). Four legit exceptions to body ≥ 18pt: chart axis labels, legends, footer / page number, and ≤ 5-word KPI sublabels (e.g. "Active users"). Descriptive sentences must be ≥ 18pt. Left-align body; center only titles and hero numbers.

**Two fonts max, one palette.** One heading font + one body font (e.g. Georgia + Calibri). One dominant brand color (60–70% weight) + one supporting + one accent. Never mix 4+ colors in body content.

**Every slide carries a non-text visual.** Shape, chart, icon, gradient band. A bullet-only deck is interchangeable with a Word doc. Exceptions: literal quote slides, code blocks, a single summary-table slide.

**Speaker notes on every content slide.** `--type notes --prop text="..."`. The speaker needs a script; the audience shouldn't read the slide verbatim.

**Preserve existing templates.** When a file already has a theme and masters, match them. Existing conventions override these guidelines.

### Visual delivery floor (applies to EVERY deck)

Before declaring done, the per-slide render (see QA) MUST satisfy:

- **Type hierarchy met** — see Requirements table above (title ≥ 36pt, body ≥ 18pt, title ≥ 2× body). Verify via `view annotated`.
- **No placeholder tokens rendered as content.** `{{name}}`, `$fy$24`, `<TODO>`, `lorem`, `xxxx`, empty `()`/`[]` in chart titles never appear.
- **No overflow past slide edges.** For 16:9 (33.87 × 19.05cm), every shape satisfies `x + width ≤ 33.87cm` AND `y + height ≤ 19.05cm`. `get` and check — don't eyeball.
- **No text overflow inside shapes.** A 72pt KPI in a 4cm-tall box clips. Shrink the number, enlarge the box, or shorten the text — never trim content to fit.
- **Cover slide is content-rich.** Title + subtitle + presenter/client block + date + a brand band or key-takeaway strap. A cover with 80% whitespace reads as a stub.
- **Contrast floor.** On dark backgrounds (brightness < 30%), body text MUST be `FFFFFF` or > 80%-bright. Mid-gray on dark navy is invisible on projection.
- **Animation restraint.** ≤ 1 per slide, ≤ 600ms — full rules in Animation restraint below.
- **No `\$`, `\t`, `\n` literals in slide text.** If `view text` shows these, a shell-escape leaked — delete and re-enter via heredoc batch.

If any fails, STOP and fix before declaring done.

### Hard rules — typography / contrast / notes / KPI fit

- **Body floors at 18pt** (absolute floor 16pt for tiny sublabels; 18pt is the working floor for all paragraph and card body text per the Requirements table). Exceptions are non-primary-read elements only: chart axis labels, legends, footer / page number, and KPI sublabels of ≤ 5 words (e.g. "Active users", "MoM growth"). Full descriptive sentences never qualify — shrink the sentence or split the slide, do not shrink the font. "The cards won't fit" is never a reason to drop below floor; it is a reason to drop cards.
- **Dark backgrounds force near-white body.** When fill brightness < 30% (`1E2761`, `36454F`, `000000`, deep forest / berry / cherry), every run of body text, card body, chart series fill, and icon color must be `FFFFFF` or brightness > 80%. Mid-gray (`6B7B8D` ≈ 44%) reads fine on a laptop screen and disappears under projector glare. Check with `view html` after the dark-fill pass.
- **Content slides carry speaker notes.** Every slide that is neither cover nor closing must have `--type notes --prop text="..."`. The speaker needs a script; the audience should not read the slide verbatim. Missing notes on a content slide is not shippable.
- **KPI text fits the card — pre-compute, don't eyeball.** In a 7cm-wide card at 60pt Georgia bold, values with `$` and `.` (wide glyphs) wrap at 4 characters. `$9.4M` breaks the card; use `$9M` + "USD millions" sublabel, or move to the 3-card 9.78cm layout. Upper bound: `max_size_pt ≈ card_width_cm × denom`, where denom = 10 for 1–2 chars, 7 for 3–4 chars, 5 for 5+ chars.

### `layout=blank` and alt text

- **`layout=blank` is the default for custom designs.** Titles become plain `shape` elements, not placeholders. `view outline` / `view issues` reporting `(untitled)` / `Slide has no title` is **expected**, not a defect. Use `layout=title` + `placeholder[title]` only when screen-reader outline compatibility matters.
- **Alt text verification.** `view stats "Pictures without alt text: 0"` is a false-positive zero (alt auto-fills to filename) — verify via `view annotated`.

## Design Principles

A deck is not a document. The audience has 3 seconds to get each slide. Before adding anything, ask: "If the audience reads only the biggest element and glances once, do they get the point?" If they have to read the bullets, the biggest element is wrong.

### Grid, margins, negative space

Standard widescreen is **33.87 × 19.05cm**. Treat it as a 12-column grid internally:

- **Edge margin ≥ 1.27cm** (0.5") on all sides.
- **Inter-block gap ≥ 0.76cm** (0.3") between cards / columns / rows.
- **≥ 20% negative space per slide.** Filling every pixel reads as amateur.
- For card grids: `usable = 33.87 − 2·margin − (N−1)·gap`, then `col_width = usable / N`. Don't hand-pick x coordinates.

### Font pairings

Two fonts max — one for headings, one for body. Pair by document register, not by novelty. "Best For" is a prompt, not a decree; if the topic matches a row, use it as the default and move on.

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

Set both fonts explicitly on every shape (`--prop font=Georgia` on title shapes, `--prop font=Calibri` on body shapes) — theme-default inheritance drifts between masters.

### Color and contrast

One dominant color does 60–70% of visual weight, two supporting tones, one accent used sparingly. Never use 4+ colors in body content. Columns are: **Primary** (dominant — the one color you see first), **Secondary** (the supporting tone), **Accent** (sparing, one-hit emphasis), **Text** (body on light fills), **Muted** (captions / axis labels / footer).

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

### Animation restraint

Each animation is a cognitive interrupt. Limits:
- **≤ 1 animation per slide**, duration **≤ 600ms**.
- Use only `fade`, `appear`, or a single `zoom-entrance` on a hero slide.
- Never: `bounce`, `swivel`, `fly-from-edge`, `spin`, multi-object choreography.
- Animation is runtime-only — verify in a live presentation viewer.

### Layout patterns & data display

Vary layout across slides — repeating the same pattern makes every slide feel identical. Pick one per slide from these building blocks:

| Pattern | When to use | Key measurement |
|---|---|---|
| **Two-column** (text left, visual right) | Concept + evidence; feature + screenshot | Each col ≈ 14-15cm; gap 1cm |
| **Icon rows** (icon in filled circle + bold header + description) | Feature lists, benefits, team roles | Icon circle 1.5-2cm; 3-4 rows max |
| **2×2 or 2×3 grid** (card tiles) | Quadrant analysis, SWOT, option comparison | Gap ≥ 0.76cm; consistent card height |
| **Half-bleed image** (full left or right half, content overlay on other side) | Hero moments, case study openers | Image 16-17cm wide; content column ≥ 14cm |
| **Large stat callout** (60-72pt number + ≤5-word sublabel below) | Single KPI, milestone, market size | Use shape, NOT a chart; sublabel 14-16pt muted |

**Data display quick rules:**
- One big number reads faster than a chart — use a `shape` with 60-72pt bold for a single KPI.
- Comparison columns (before/after, A vs B) beat a table for 2-3 options.
- Timelines and process flows: numbered step shapes + connectors, not a bullet list.

### Visual motif commitment

Pick ONE distinctive element (rounded image frames, section numbers in filled circles, single-side border band, diagonal accent strips) and carry it to every slide. Declare it in your build plan first: `## Motif: numbered circles in brand color`.

### What to avoid (common design mistakes)

These are the patterns that make a deck look AI-generated or amateur:

- **NEVER place a decorative line under slide titles.** Underline stripes below headings are the single most common AI-slide tell. Use whitespace or background color change instead.
- **Don't repeat the same layout across consecutive slides.** Alternate between two-column, callout, grid, and half-bleed patterns. Same layout = same visual rhythm = audience tunes out.
- **Don't center body text.** Left-align all paragraphs, lists, card descriptions. Center only slide titles and hero numbers.
- **Don't default to blue** because it feels "professional." Pick the palette that fits the topic — finance reads navy, sustainability reads forest, energy reads coral.
- **Don't use inconsistent spacing.** Choose either 0.76cm or 1.27cm as your inter-block gap and use it everywhere. Mixed gaps look unfinished.
- **Don't create text-only slides.** If a slide has only a title and bullets, add a supporting shape, chart, icon, or image. A purely textual slide is a Word paragraph.
- **Don't style one slide and leave the rest plain.** Commit fully or keep it simple throughout — partial styling reads as abandoned.

## Common Workflow

1. **Open/close mode.** Always `officecli open <file>` at start + `officecli close <file>` at end. Resident is the default, not an optimization. Use `batch` in ≤ 12-op chunks for repetitive shape grids.
2. **Orient.** New deck: `officecli create "$FILE"`. Existing: `officecli view "$FILE" outline` first. Never edit blind.
3. **Build in display order — HARD RULE.** `--index` on slide add is frequently ignored. Add slides in audience-view order: cover → agenda → section-1 divider → section-1 content → section-2 divider → … → closing. Out-of-order insertion requires `officecli move "$FILE" /slide[N] --index M` + re-verify with `get --depth 0`. **Before final delivery, confirm slide count + narrative arc match your build plan.** Gate 3's order-sanity check catches cases where the cover ends up as slide 11 of 14 instead of slide 1.
4. **Incremental per slide.** Create slide + background, then title, then supporting shapes / charts / connectors. Always `layout=blank` for custom designs. After each structural op, `get /slide[N] --depth 1` to confirm shape IDs.
5. **Format to spec.** Per the Requirements table; formatting is deliverable, not polish.
6. **Close + verify.** `officecli close` writes the ZIP. Always open in the target presentation viewer before shipping — chart colors, animations, fonts, and zoom are runtime features `view html` can't render. Full verification in QA below.
7. **QA — assume there are problems.** Fix-and-verify until a cycle finds zero new issues.

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

officecli close "$FILE"
officecli validate "$FILE"
```

Shape of every build: open → slide+background → title → body → notes → close → validate.

## Reading & Analysis

Start wide, then narrow. `outline` first, `view text` / `get` / `query` once you know where to look.

```bash
officecli view "$FILE" outline          # slide count, titles, shape counts (undercounts tables/charts)
officecli view "$FILE" annotated        # complete per-slide breakdown with fonts, sizes, tables, charts
officecli view "$FILE" text --start 1 --end 5   # text dump (does NOT extract table cells — use get)
officecli view "$FILE" issues           # empty slides, overflow hints
officecli view "$FILE" stats            # counts + missing alt (false-positive zero — verify via view annotated)
```

**Inspect one element.** XPath-style paths, 1-based. ALWAYS quote.

```bash
officecli get "$FILE" "/slide[1]" --depth 1              # shape list with IDs and names
officecli get "$FILE" "/slide[1]/shape[@name=Title]"     # @name selector (LEAD — stable across reorderings)
officecli get "$FILE" "/slide[1]/shape[@id=10007]"       # @id selector (also stable)
officecli get "$FILE" "/slide[1]/chart[1]"               # chart data + config
officecli get "$FILE" "/slide[1]/table[1]" --depth 3     # table rows / cells
```

Add `--json` for machine output. Use `[last()]` for "the last": `/slide[last()]/shape[1]`.

**Query across the deck.** CSS-like selectors; operators `=`, `!=`, `~=`, `>=`, `<=`, `[attr]`:

```bash
officecli query "$FILE" 'shape:contains("Revenue")'
officecli query "$FILE" 'picture:no-alt'                 # accessibility gap
officecli query "$FILE" 'shape[fill=1E2761]'             # color match
officecli query "$FILE" 'shape[width>=10cm]'             # numeric
officecli query "$FILE" 'animation'                      # every animation in the deck
```

**`query --json` output schema.** Results wrap in `.data.results[]`. To extract the first result's ID: `jq -r '.data.results[0].format.id'` — NOT `.[0].id`. Shape name is `.name`; fill is `.format.fill`; textColor is `.format.textColor`. Verify with `query ... --json | jq .data.results[0]` on an unknown shape.

**Visual preview (LEAD).**

```bash
officecli view "$FILE" html                # prints an HTML preview path; Read it for per-slide visual audit (best structural ground truth)
officecli view "$FILE" svg --start 3 --end 3   # single slide SVG (charts + gradients do NOT render in SVG)
```

## Creating & Editing

Verbs: `add` / `set` / `remove` / `move` / `swap` / `batch` / `raw-set`. Ninety percent of a deck is slides, shapes, text, a few charts, pictures, connectors.

### Slides and backgrounds

A slide is `/slide[N]`. Always pass `layout=blank` for custom designs. Background: solid, gradient, or image.

```bash
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761                 # solid
officecli add "$FILE" / --type slide --prop layout=blank --prop "background=1E2761-CADCFC-180"   # gradient (start-end-angle)
officecli add "$FILE" / --type slide --prop layout=blank --prop "background.image=hero.jpg"      # image background (LEAD)
```

### Shapes

A `shape` holds text, fill, border, position, and optional animation / link.

```bash
officecli add "$FILE" /slide[2] --type shape --prop name=Title --prop text="Key Insight" \
  --prop x=2cm --prop y=2cm --prop width=20cm --prop height=3cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
```

Positioning is explicit — no layout engine, you own the grid math. `--prop preset=` picks geometry (`rect`, `roundRect`, `ellipse`, `triangle`, `arrow`, `star5`, ...); custom `M...Z` paths are NOT supported at the high-level prop layer (use preset or raw-set `a:custGeom`). **Name shapes at creation** (`--prop name=HeroTitle`) and address later with `"/slide[N]/shape[@name=HeroTitle]"` — positional `/shape[3]` breaks after any z-order / remove.

> **ID semantics.** IDs are assigned per-XML-element, not per-`add`-command. Paragraphs and runs consume IDs too — so the 4 IDs returned by 4 `add shape` calls are NOT guaranteed to be sequential (child paragraphs ate some). After a rebuild or remove-then-add, re-`get --depth 1` before referencing IDs. **Prefer `@name=` over `@id=`** — names are stable across all structural ops.

### Text inside shapes (paragraphs, runs, styling)

A shape has paragraphs (`paragraph[K]`) and runs. For one-line text, `--prop text=` on the shape is enough. Multi-line or mixed styling:

```bash
# add --type paragraph accepts only text + align; styling goes through a follow-up set or an add --type run:
officecli add "$FILE" "/slide[2]/shape[@name=Card1]" --type paragraph --prop text="First bullet"
officecli set "$FILE" "/slide[2]/shape[@name=Card1]/paragraph[1]" --prop bold=true --prop size=20 --prop color=FFFFFF

# Styled run in one step:
officecli add "$FILE" "/slide[2]/shape[@name=Card1]/paragraph[1]" --type run \
  --prop text=" (inline detail)" --prop size=14 --prop italic=true --prop color=8899BB
```

For real newlines inside one run, use a batch heredoc with JSON `"\n"`. Shell-quoted `\n` in `--prop text=` is NOT interpreted.

### Charts

Chart types: column, bar, line, pie, doughnut, scatter, area, waterfall, funnel, boxWhisker. Pick per the Design Principles chart-choice table. Two data forms:

```bash
# (a) compact inline — quick demo charts
officecli add "$FILE" /slide[3] --type chart --prop chartType=column \
  --prop "data=Revenue:42,45,48;Growth:2,7,7" --prop "categories=Q1,Q2,Q3" \
  --prop x=2cm --prop y=4cm --prop width=20cm --prop height=10cm --prop title="FY26"

# (b) dotted per-series — multi-series with explicit brand colors (typical case)
officecli add "$FILE" /slide[3] --type chart --prop chartType=column \
  --prop series1.name=Revenue --prop series1.values="42,45,48" --prop series1.color=1E2761 \
  --prop series2.name=Growth  --prop series2.values="2,7,7"    --prop series2.color=CADCFC \
  --prop categories="Q1,Q2,Q3" \
  --prop x=2cm --prop y=4cm --prop width=20cm --prop height=10cm
```

Gotchas: (1) series cannot be added after creation — include all series at `add` time or `remove` + re-add. (2) chart titles with `()`, `[]`, `TBD` ship as literal text — replace before delivery. (3) some viewers normalize chart colors to theme defaults — verify in the target presentation viewer.

### Pictures

```bash
officecli add "$FILE" /slide[4] --type picture --prop src=hero.jpg \
  --prop x=1cm --prop y=1cm --prop width=32cm --prop height=18cm \
  --prop alt="Product hero, gradient lit from right"
```

Confirm with `officecli query "$FILE" 'picture:no-alt'` — must be empty before delivery (but remember `view stats` is a false-positive zero because alt auto-fills to filename).

### Connectors (LEAD — flowcharts / decision trees first-class)

Draws a line between two shapes or free coordinates. CLI-native props: `shape`, `from`, `to`, `x`, `y`, `width`, `height`, `color`, `headEnd`, `tailEnd` (values: `triangle|arrow|stealth|diamond|oval|none`). `line=`, `lineWidth=`, `lineDash=` are UNSUPPORTED — use raw-set `a:ln` for custom line styling.

- `shape` enum: short form `straight | elbow | curve`, or storage form `straightConnector1 | bentConnector3 | curvedConnector3 | line`. `bentConnector2` / `curvedConnector2` are rejected.
- `from=`/`to=` accept the same shape-ref forms as the rest of the CLI: bare integer (shape ID), `/slide[N]/shape[M]` (positional), `/slide[N]/shape[@id=M]`, or `/slide[N]/shape[@name=Foo]`.
- Arrowheads via `--prop tailEnd=triangle` (or `headEnd=` for reverse direction) — **requires CLI 1.0.63+**. On older 1.0.60–1.0.62 the `tailEnd=` / `headEnd=` props were UNSUPPORTED on connector; fall back to raw-set `<a:tailEnd type="triangle" w="med" len="med"/>` on `/connector[@id=ID]/spPr/ln`. Accepted values: `triangle | arrow | stealth | diamond | oval | none` (plus `closed`/`open`/`circle`, parsed by `ParseLineEndType`). For custom arrow size `w`/`len` on any version, use raw-set.

```bash
officecli add "$FILE" /slide[5] --type connector \
  --prop "from=/slide[5]/shape[@name=BoxA]" --prop "to=/slide[5]/shape[@name=BoxB]" \
  --prop shape=elbow --prop color=333333 --prop tailEnd=triangle

# Optional — raw-set for custom arrow size (w=med, len=med):
# CONN_ID=$(officecli query "$FILE" 'connector' --json | jq -r '.data.results[-1].format.id')
# officecli raw-set "$FILE" "/slide[5]/connector[@id=$CONN_ID]/spPr/ln" \
#   --action append --xml '<a:tailEnd type="triangle" w="med" len="med"/>'
```

**Every flow connector needs an arrowhead.** Without one, `bentConnector3` renders as a directionless line. Use `--prop tailEnd=triangle` on the connector add or set. `preset=rightArrow` overlay only works for horizontal flows; diamonds / decision trees with diverging edges need `tailEnd=`.

### Animations (LEAD)

One preset per slide, ≤ 600ms. Set via shape-level prop or via deep-path `add --type animation`:

```bash
officecli set "$FILE" "/slide[2]/shape[@name=HeroCard]" --prop animation=fade-entrance-400
officecli get "$FILE" "/slide[2]/shape[@name=HeroCard]" --json | jq .animation
officecli set "$FILE" "/slide[2]/shape[@name=HeroCard]" --prop animation=none    # remove all animations on shape
officecli remove "$FILE" "/slide[2]/shape[@name=HeroCard]/animation[1]"          # remove a specific animation by index
```

**Get round-trip (1.0.58+).** `get animation[N]` now returns the `trigger` field as well — `onClick | afterPrevious | withPrevious` — so Add/Set and Get round-trip. Verify with `officecli get "$FILE" "/slide[N]/shape[@name=X]/animation[1]" --json | jq '.trigger,.duration'` if you need to confirm the read-back matches what you set.

### Hyperlinks, tooltips, slide-jump

```bash
officecli set "$FILE" "/slide[7]/shape[@name=NavBtn]" --prop link=slide:2 --prop tooltip="Back to Agenda"
officecli set "$FILE" "/slide[7]/shape[@name=DocsBtn]" --prop link=https://example.com
```

### Tables, placeholders, groups, zoom — one-liners

- **Tables** — `--type table --prop rows=N --prop cols=M`. Row-level `set` supports `height`, `header`, `c1/c2/c3`. Cell formatting lives on the cell paragraph / run. Populate rows BEFORE setting table-level font (font cascade gets reset by row ops).
- **Placeholders** — `"/slide[N]/placeholder[title]"` / `placeholder[body]`. Available only when the slide uses a layout with placeholders (not `layout=blank`).
- **Groups** (LEAD) — address children via `"/slide[N]/group[@name=G]/shape[1]"`. Survives reordering better than positional indexes.
- **Zoom slide** (LEAD) — `--type zoom --prop targets="3,7,15"`. Section-navigation hub. Zoom is a runtime feature — `view html` shows the static geometry; the zoom interaction runs only in a live presentation viewer.
- **Slide comments** — reviewer annotations anchored at `/slide[N]/comment[M]`. Full lifecycle (`add / set / get / query / remove`). Props: `text`, `author`, `initials` (auto-derived), `date` (ISO 8601, defaults to UtcNow), `x` / `y` (EMU anchor).
  ```bash
  officecli add "$FILE" "/slide[2]" --type comment --prop author="Alice" --prop text="Tighten this bullet" --prop x=20cm --prop y=3cm
  officecli query "$FILE" 'comment' --json | jq '.data.results | length'   # count all review comments
  officecli remove "$FILE" "/slide[2]/comment[1]"                           # resolve after addressing
  ```

### Deck-level recipes

Six patterns that aren't obvious from the primitives: (a) cover + section divider, (b) chart + commentary, (c) 4-step flowchart, (d) 10-slide board-review skeleton, (d′) 20-slide Series B / pitch-deck skeleton, (e) KPI callouts, (f) YES/NO decision tree. Each describes the **visual outcome** first, then a runnable block. `$FILE` stands in for your filename.

#### (a) Executive cover + section divider

**Visual outcome.** Cover: dark navy fill, centered 44pt title, 18pt ice-blue meta line. Section divider: same dark fill, large translucent "02" (120pt, 15% opacity) behind a 40pt title — the number becomes a background graphic, the title carries the message.

```bash
# Cover
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761
officecli add "$FILE" "/slide[last()]" --type shape --prop text="Strategic Growth Review" \
  --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=3cm \
  --prop font=Georgia --prop size=44 --prop bold=true --prop color=FFFFFF --prop align=center
officecli add "$FILE" "/slide[last()]" --type shape --prop text="Prepared for Acme Leadership — FY26 Outlook" \
  --prop x=2cm --prop y=11cm --prop width=29.87cm --prop height=1.2cm \
  --prop font=Calibri --prop size=18 --prop color=CADCFC --prop align=center

# Section divider — translucent number added FIRST (stays behind), title last (stays on top)
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761
officecli add "$FILE" "/slide[last()]" --type shape --prop text="02" \
  --prop x=2cm --prop y=3cm --prop width=29.87cm --prop height=10cm \
  --prop font=Georgia --prop size=120 --prop bold=true --prop color=FFFFFF --prop opacity=0.15 --prop align=center
officecli add "$FILE" "/slide[last()]" --type shape --prop text="Financial Performance" \
  --prop x=2cm --prop y=7.5cm --prop width=29.87cm --prop height=2.5cm \
  --prop font=Georgia --prop size=40 --prop bold=true --prop color=FFFFFF --prop align=center
```

**Z-order.** Later-added shapes are on top. Add background decoration FIRST, titles LAST. If the order got flipped, fix with `--prop zorder=back/front` — but that renumbers siblings, so re-`get --depth 1` before stacking more.

#### (b) Data slide (chart + commentary block)

**Visual outcome.** Left two-thirds: column chart with brand series colors. Right one-third: a "Key Insight" card with 20pt heading and 18pt body — the audience reads the takeaway before parsing the bars.

```bash
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[last()]" --type shape --prop text="FY26 Revenue Beat Plan by 18%" \
  --prop x=1.5cm --prop y=1cm --prop width=30cm --prop height=1.8cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none

# Chart — left 2/3
officecli add "$FILE" "/slide[last()]" --type chart --prop chartType=column \
  --prop series1.name=Actual --prop series1.values="42,45,48,55" --prop series1.color=1E2761 \
  --prop series2.name=Plan   --prop series2.values="40,42,45,48" --prop series2.color=CADCFC \
  --prop categories="Q1,Q2,Q3,Q4" \
  --prop x=1.5cm --prop y=3.5cm --prop width=20cm --prop height=14cm \
  --prop title='FY26 Revenue ($M)'

# Commentary card — right 1/3 (background card + heading + body)
officecli add "$FILE" "/slide[last()]" --type shape --prop preset=roundRect --prop fill=F5F7FA --prop line=none \
  --prop x=22.5cm --prop y=3.5cm --prop width=9.8cm --prop height=14cm
officecli add "$FILE" "/slide[last()]" --type shape --prop text="Key Insight" \
  --prop x=23cm --prop y=4cm --prop width=9cm --prop height=1.2cm \
  --prop font=Georgia --prop size=20 --prop bold=true --prop color=1E2761 --prop fill=none
officecli add "$FILE" "/slide[last()]" --type shape --prop text="EMEA launch + NRR at 118% drove 12pp of the 18pp beat." \
  --prop x=23cm --prop y=5.5cm --prop width=9cm --prop height=11cm \
  --prop font=Calibri --prop size=18 --prop color=333333 --prop fill=none

officecli add "$FILE" "/slide[last()]" --type notes --prop text="Lead with the 18% beat, then the EMEA + NRR story."
```

`$` in chart title: wrap the whole prop in single quotes (`--prop title='FY26 Revenue ($M)'`) to prevent shell expansion.

#### (c) Flowchart / process diagram (connectors + shapes) — batch-heredoc

**Visual outcome.** Four rounded boxes across the slide at y=8cm, each 6cm × 3cm, connected left-to-right with elbow connectors + triangle arrowheads. Boxes alternate dominant / supporting fill. 32pt title above.

Grid math for 4 boxes across a 33.87cm slide with 1.5cm margins: `gap = (33.87 − 2·1.5 − 4·6) / 3 = 2.29cm`. Box x-positions: `1.5, 9.79, 18.08, 26.37`. **Batch heredoc is portable** (no `bc`, no bash arrays); pre-compute coordinates in the JSON.

```bash
# Use explicit slide index — [last()] is rejected in some resident versions (see Pitfalls).
# COUNT slides via query (works on closed AND resident-open files; get --depth 0 default output is not JSON).
N_SLIDE=$(officecli query "$FILE" 'slide' --json | jq '.data.results | length')
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
SLIDE=$((N_SLIDE + 1))

officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop name=FlowTitle \
  --prop text="Onboarding in 4 Steps" \
  --prop x=1.5cm --prop y=1cm --prop width=30cm --prop height=1.8cm \
  --prop font=Georgia --prop size=32 --prop bold=true --prop color=1E2761 --prop fill=none

# 4 boxes + overlaid labels in one batch (alternating navy/iceblue fills).
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step1","preset":"roundRect","fill":"1E2761","line":"none","x":"1.5cm","y":"8cm","width":"6cm","height":"3cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step2","preset":"roundRect","fill":"CADCFC","line":"none","x":"9.79cm","y":"8cm","width":"6cm","height":"3cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step3","preset":"roundRect","fill":"1E2761","line":"none","x":"18.08cm","y":"8cm","width":"6cm","height":"3cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step4","preset":"roundRect","fill":"CADCFC","line":"none","x":"26.37cm","y":"8cm","width":"6cm","height":"3cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Step 1","x":"1.5cm","y":"8.5cm","width":"6cm","height":"1cm","font":"Georgia","size":"20","bold":"true","color":"FFFFFF","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Step 2","x":"9.79cm","y":"8.5cm","width":"6cm","height":"1cm","font":"Georgia","size":"20","bold":"true","color":"1E2761","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Step 3","x":"18.08cm","y":"8.5cm","width":"6cm","height":"1cm","font":"Georgia","size":"20","bold":"true","color":"FFFFFF","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Step 4","x":"26.37cm","y":"8.5cm","width":"6cm","height":"1cm","font":"Georgia","size":"20","bold":"true","color":"1E2761","align":"center","fill":"none"}}
]
EOF

# Connectors. Arrowhead via --prop tailEnd=triangle on add (CLI-native).
for pair in "Step1 Step2" "Step2 Step3" "Step3 Step4"; do
  A=$(echo $pair | cut -d' ' -f1); B=$(echo $pair | cut -d' ' -f2)
  officecli add "$FILE" "/slide[$SLIDE]" --type connector \
    --prop "from=/slide[$SLIDE]/shape[@name=$A]" \
    --prop "to=/slide[$SLIDE]/shape[@name=$B]" \
    --prop shape=elbow --prop color=333333 --prop tailEnd=triangle
done
```

`shape=elbow` is the canonical short form. `bentConnector3` also works as storage name; `bentConnector2` is **rejected**. `query --json` wraps results in `.data.results[]` — `.[0].id` is WRONG, use `.data.results[0].format.id`.

#### (d) Quarterly review deck skeleton (10-slide blueprint)

**Visual outcome.** A board-ready narrative arc in ten slides. No copy-paste block — the structure below is the decision tree; fill each slide using recipes (a)–(c) and the Design Principles.

| # | Slide | Layout | Recipe | Notes |
|---|---|---|---|---|
| 1 | Cover | full dark fill | (a) cover | Title + meta line + brand band |
| 2 | Agenda | white fill | numbered list | 5 section titles, 24pt, left-aligned; match exactly what follows |
| 3 | Executive summary | white fill | 3 KPI callouts | Three big numbers (60pt) across the top, one-line qualifier under each, 2-line narrative at the bottom |
| 4 | Section divider: Financials | dark | (a) divider | Big "01" + "Financial Performance" |
| 5 | Revenue vs plan | white | (b) chart + commentary | Column chart, insight card |
| 6 | Margin walk | white | (b) chart + commentary | Waterfall chart, key drivers list |
| 7 | Section divider: Growth | dark | (a) divider | Big "02" + "Growth Initiatives" |
| 8 | GTM motion | white | (c) flowchart | 4-step process, connectors |
| 9 | Roadmap timeline | white | timeline shapes + connectors | 4 quarters as circles on a line, one deliverable under each |
| 10 | Thank you / next steps | dark | (a) cover variant | One bullet per next step, max 3 bullets |

**Build it.** `officecli open "$FILE"` → loop the slides with the appropriate recipe → `officecli close "$FILE"` → run the Delivery Gate → spot-check in PowerPoint. A 10-slide deck assembled this way takes ~30–50 commands; use `batch` heredocs in ≤ 12-op chunks for the repetitive shapes.

#### (d′) Series B / pitch deck skeleton (20-slide blueprint)

**Visual outcome.** A VC-ready narrative in twenty slides. Same structure as (d) but calibrated for fundraising rhythm: problem → solution → market → product → model → traction → team → financials → ask → close.

| # | Slide | Layout | Recipe | Notes |
|---|---|---|---|---|
| 1 | Cover | dark fill | (a) cover | Company name + one-line tagline + round/amount + date |
| 2 | TL;DR / The Ask | dark | 3 KPI callouts | "$35M Series B · 72% GM · 4.2x LTV:CAC" style |
| 3 | Section 01 — Problem | dark | (a) divider | Big "01" + one-line problem statement |
| 4 | Problem data | white | (b) chart + commentary | 3 data cards + sourced footnote, "industry sells X, customers want Y" |
| 5 | Section 02 — Solution | dark | (a) divider | Big "02" + product tagline |
| 6 | How it works | white | (c) flowchart | 4-step product loop, connectors with arrowheads |
| 7 | Section 03 — Market | dark | (a) divider | Big "03" + market size one-liner |
| 8 | TAM / SAM / SOM | white | (b) chart + commentary | Column chart + "Why now" list (macro drivers) |
| 9 | Section 04 — Product | dark | (a) divider | Big "04" + product positioning |
| 10 | Product in three pieces | white | 3-card row | Hardware / software / marketplace, with prices |
| 11 | Section 05 — Traction | dark | (a) divider | Big "05" + "We shipped. People stayed." |
| 12 | ARR trajectory | white | (b) chart + commentary | Line chart + callout number |
| 13 | Retention cohort | white | (b) chart + commentary | Cohort chart + NPS / App Store / referral stats |
| 14 | Section 06 — Business model | dark | (a) divider | Big "06" + unit-econ summary |
| 15 | Unit economics | white | 4 KPI callouts | CAC / LTV / GM / payback — the VC napkin slide |
| 16 | Section 07 — Team | dark | (a) divider | Big "07" + team one-liner |
| 17 | Founders + advisors | white | 4-card grid + advisory row | Real prior companies, real names |
| 18 | Section 08 — Financials | dark | (a) divider | Big "08" + trajectory tagline |
| 19 | 4-year plan | white | (b) chart + commentary | Hockey stick + honest assumptions panel |
| 20 | The Ask / Thank you | dark | (a) cover variant | `$XX M` hero number + 3 bullet use-of-funds + contact |

Parallel to (d) — swap recipes per row; each divider must appear BEFORE its section content (see Gate 3 order sanity).

#### (e) KPI callouts — giant-number card grid

**Visual outcome.** Three (or four) giant numbers across a row, each with a one-line unit sublabel + small percent-change chip + one-line takeaway underneath. This is the single most common exec-deck element.

**KPI value sizing.** 60pt Georgia bold fits ~5 chars in a 9.78cm card (`$84.2`, `118%`, `24.5`). For values that don't fit (`$84.2M`), split: `$84.2` as the big number, `USD millions` as the sublabel — never shrink the font to chase the unit suffix, it just wraps.

Grid math for 3 cards across, 1.5cm margins, 0.76cm gap: `col_width = (33.87 − 3 − 2·0.76) / 3 = 9.78cm`. x-positions: `1.5, 12.04, 22.58`.

```bash
# 3 KPI cards (navy fill, white number, ice-blue sublabel, terracotta pct chip for the watch-card).
# All 12 adds in one batch. Replace $SLIDE with your target slide index.
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"K1Card","preset":"roundRect","fill":"1E2761","line":"none","x":"1.5cm","y":"4cm","width":"9.78cm","height":"7cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"84.2","x":"1.5cm","y":"4.8cm","width":"9.78cm","height":"2.8cm","font":"Georgia","size":"60","bold":"true","color":"FFFFFF","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"USD millions · ARR","x":"1.5cm","y":"8cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","color":"CADCFC","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"+24% YoY","x":"1.5cm","y":"9cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","bold":"true","color":"CADCFC","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"K2Card","preset":"roundRect","fill":"1E2761","line":"none","x":"12.04cm","y":"4cm","width":"9.78cm","height":"7cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"118%","x":"12.04cm","y":"4.8cm","width":"9.78cm","height":"2.8cm","font":"Georgia","size":"60","bold":"true","color":"FFFFFF","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Net revenue retention","x":"12.04cm","y":"8cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","color":"CADCFC","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"best-in-class floor 115%","x":"12.04cm","y":"9cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","italic":"true","color":"CADCFC","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"K3Card","preset":"roundRect","fill":"B85042","line":"none","x":"22.58cm","y":"4cm","width":"9.78cm","height":"7cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"$1.42","x":"22.58cm","y":"4.8cm","width":"9.78cm","height":"2.8cm","font":"Georgia","size":"60","bold":"true","color":"FFFFFF","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"CAC payback (yrs)","x":"22.58cm","y":"8cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","color":"FFFFFF","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"+8% — watch","x":"22.58cm","y":"9cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","bold":"true","color":"FFFFFF","align":"center","fill":"none"}}
]
EOF
```

Use the accent color (terracotta here) on the single "watch" card so the audience reads risk in one second. Narrative headline above: "Beat plan on 2 of 3 metrics — CAC is the watch-item" beats "FY26 Q1 KPIs" every time.

#### (f) Decision tree — YES/NO branching with diverging-then-converging edges

**Visual outcome.** A decision diamond at the top; two child boxes (YES / NO path) diverging left-right; each path converges to a shared terminal box. Used for LOTO / chemical / fire safety training, compliance triage, escalation rules.

Layout: diamond at top-center (x=13.94, y=2cm, 6×3cm). YES branch at x=3cm y=7.5cm; NO branch at x=24.87cm y=7.5cm; terminal at x=13.94cm y=13cm.

```bash
# Each box carries its own text via valign=middle (no separate text-overlay shape).
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Decide","preset":"diamond","fill":"1E2761","line":"none","x":"13.94cm","y":"2cm","width":"6cm","height":"3cm","text":"Hazardous energy present?","font":"Calibri","size":"14","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"YesBox","preset":"roundRect","fill":"B85042","line":"none","x":"3cm","y":"7.5cm","width":"8cm","height":"3cm","text":"Lockout + Tagout + Verify","font":"Calibri","size":"16","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"NoBox","preset":"roundRect","fill":"CADCFC","line":"none","x":"22.87cm","y":"7.5cm","width":"8cm","height":"3cm","text":"Proceed with standard PPE","font":"Calibri","size":"16","bold":"true","color":"1E2761","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Done","preset":"roundRect","fill":"2C5F2D","line":"none","x":"13.94cm","y":"13cm","width":"6cm","height":"2.5cm","text":"Begin service","font":"Calibri","size":"16","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"YES","x":"6.5cm","y":"5.8cm","width":"2cm","height":"1cm","font":"Calibri","size":"14","bold":"true","color":"B85042","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"NO","x":"25.5cm","y":"5.8cm","width":"2cm","height":"1cm","font":"Calibri","size":"14","bold":"true","color":"1E2761","align":"center"}}
]
EOF

# 4 connectors: Decide→YesBox, Decide→NoBox, YesBox→Done, NoBox→Done.
for pair in "Decide YesBox" "Decide NoBox" "YesBox Done" "NoBox Done"; do
  A=$(echo $pair | cut -d' ' -f1); B=$(echo $pair | cut -d' ' -f2)
  officecli add "$FILE" "/slide[$SLIDE]" --type connector \
    --prop "from=/slide[$SLIDE]/shape[@name=$A]" \
    --prop "to=/slide[$SLIDE]/shape[@name=$B]" \
    --prop shape=elbow --prop color=333333 --prop tailEnd=triangle
done
```

Color convention: red path = stop/escalate, blue path = standard-action, green terminal = safe end-state. Trainees reading a decision tree backwards is a real life-safety risk — every connector needs an arrowhead.

## QA (Required)

**Assume there are problems.** First render is almost never correct. If you found zero issues, you were not looking hard enough.

### Delivery Gate (any failure = REJECT, do NOT deliver)

Three checks. Gate 1 is the schema defense; Gate 2 catches overflow / format / structure issues; Gate 3 is the only visual-assembly check. **None of Gates 1–2 can see a rendered slide.** Refuse to declare done until every gate prints its OK message.

```bash
FILE="deck.pptx"

# Gate 1 — schema check (REJECT on any validate error)
officecli validate "$FILE" && echo "Gate 1 OK" || { echo "REJECT Gate 1"; exit 1; }

# Gate 2 — overflow / format / structure issues (filter expected layout=blank "no title" noise)
ISSUES=$(officecli view "$FILE" issues 2>&1 | grep -vE "Slide has no title")
echo "$ISSUES" | grep -qE "^\s*\[[A-Z][0-9]+\]" && { echo "REJECT Gate 2:"; echo "$ISSUES"; exit 1; } || echo "Gate 2 OK"

echo "Delivery Gate 1–2 PASS — proceed to Gate 3 (fresh-eyes visual audit)"
```

### Gate 3 — Visual audit via HTML preview (MANDATORY)

Run `officecli view "$FILE" html` and Read the returned HTML path. For every slide:

- **overlap** — text shapes overlap each other / a chart, or a giant decorative number (01/02/03 at 100pt+) collides with a divider title
- **dark-on-dark** — text on fill brightness < 30% with text brightness < 80% (quiz options, phone numbers, card labels on navy/red/green)
- **missing arrowheads** — flowchart / decision-tree connectors render as plain lines
- **order sanity** — slide sequence matches the narrative (cover → agenda → dividers-before-sections → closing)

REJECT and list every instance with slide number; else report "Gate 3 PASS".

## Common Pitfalls

Sanity-check cheatsheet — what breaks on the first try. Design + shell traps.

| Pitfall | Correct approach |
|---|---|
| Unquoted `[N]` in zsh/bash | Always quote paths: `"/slide[1]"`. zsh globs unquoted `[1]` → `no matches found` — #1 first-use stumble |
| `--name "foo"` | All attributes go through `--prop`: `--prop name="foo"` |
| Guessing a prop name | `officecli help pptx <element>` — don't improvise |
| `/shape[myname]` (bare name in brackets) | Use `@name=` selector: `/shape[@name=myname]` or `/shape[@id=10007]` |
| Positional `/shape[3]` after z-order / remove | Positions drift. Use `@name=` / `@id=` for any repeated reference |
| `[last]` without parens | Must be `[last()]`: `/slide[last()]/shape[1]` |
| `/slide[last()]` in resident mode | Some resident versions reject it with "Shapes must be added to a slide: /slide[N]". Use explicit `/slide[N]` from `get --depth 0` for production builds. |
| Paths 1-based vs `--index` 0-based | `/slide[1]` = first slide; `--index 0` = first position |
| `$` in `--prop text=` | Single-quote: `--prop text='$15M'`. Double-quoted `"$15M"` gets shell-expanded to `M` |
| `\n` / `\t` in `--prop text=` | CLI does NOT interpret. Use multiple `--type paragraph`, or batch heredoc with JSON `"\n"` |
| Modifying a file open in PowerPoint | Close it in PowerPoint first |
| `view issues "Slide has no title"` on `layout=blank` | Expected on blank layouts (titles are shapes, not placeholders). Not a defect |
