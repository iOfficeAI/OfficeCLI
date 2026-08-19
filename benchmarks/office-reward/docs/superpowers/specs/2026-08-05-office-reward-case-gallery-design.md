# Office Reward Case Gallery Design

## Goal

Extend the completed Office Reward progress site with a comparison-oriented
gallery of nine complete scoring cases while preserving the current visual
system and Cloudflare Quick Tunnel URL.

## Case Set

The gallery contains nine cases, three per format:

- PowerPoint: executive review, research result slide, product roadmap;
- Word: quarterly report, table-heavy analysis, structured policy document;
- Excel: revenue model, chart dashboard, formula-driven operating sheet.

Every case uses a distinct real rendered asset. Case data follows the actual
Office reward result contract and includes:

- document format and case title;
- unit type and unit count;
- Aesthetics, Content Accuracy, and Communication Effectiveness scores;
- `overall_raw_score_100` and `reward_0_1`;
- `coverage_0_1` and `complete` status;
- short deterministic and model-evidence summaries.

The displayed cases are deterministic showcase fixtures, not live API calls.
Their values must remain internally consistent with the reward formula.

## Interface

Add a new `#cases` section after the format overview.

- Navigation gains a `Cases` link.
- A segmented filter switches between All, PPT, Word, and Excel.
- Desktop uses a three-column dense score grid; mobile uses one column.
- Each card contains a real thumbnail, format marker, title, unit summary,
  three compact score bars, reward, coverage, and status.
- Selecting a card opens an accessible detail dialog with the complete score
  breakdown, evidence, issue summary, and reward calculation.
- Existing status, pipeline, test, evidence, and delivery sections remain.

Cards use the existing restrained palette, square geometry, and border system.
No nested cards, decorative gradients, or marketing copy are added.

## Assets

Create nine unique raster assets:

- PPT assets may use suitable real slide images already tracked in the scoring
  repository;
- Word and Excel assets are generated from real OfficeCLI documents and
  rendered through the existing Office tooling;
- all delivered assets are local files under `public/cases/`.

The site must not depend on remote image URLs.

## Implementation

Keep `app/page.tsx` server-rendered. Add a small client component for filtering
and the detail dialog. Case fixtures live in a typed local module so score
consistency can be tested without network access.

## Verification

- validate all nine assets exist and are nonblank;
- verify every case reward equals raw overall divided by 100;
- verify filters and dialog labels are present in rendered output;
- run the site build and rendered HTML tests;
- verify the existing local URL and current Cloudflare Quick Tunnel return the
  new case gallery.

## Deployment

The existing Quick Tunnel continues to target `http://localhost:3000`.
Vite already allows `.trycloudflare.com`, so the public URL updates when the
local development server reloads. The tunnel remains temporary and has no
uptime guarantee.
