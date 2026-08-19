# Office Reward Scoring Workbench

PPTX, DOCX, and XLSX fine-grained evaluation experiments plus a browser-based
human annotation workbench. The UI puts the rubric, rendered Office unit,
direct model score, and blank human score side by side. Human entries stay in
the browser's `localStorage` and never overwrite model evidence.

## Current Evidence

| Item | Value |
| --- | ---: |
| Office units | 54 |
| PPT / Word / Excel | 30 / 12 / 12 |
| Represented multimodal calls | 108 |
| Represented model tokens | 479,688 |
| Direct scored subquestions | 2,160 |
| Explicit Content Accuracy abstentions | 1,080 |

Each Aesthetics and Communication Effectiveness dimension contains five parent
rubric blocks. Every parent is decomposed into four directly scored
subquestions; the parent score is the arithmetic mean of those four children.
Content Accuracy remains `N/A` when no authoritative reference is attached.

The 30-slide PPT sample has historical coarse human aesthetics grades. Its
model-versus-human Spearman correlation is `0.466303`, with visible score
inflation. Word and Excel do not yet have human labels. These results are
diagnostic and do not establish human alignment.

## Run The Site

From the OfficeCLI repository root, use Node.js `>=22.13.0`:

```bash
cd benchmarks/office-reward
npm ci
npm run dev
```

Open `/rubric` for the four-region annotation workbench. The root route contains
the Office reward mechanism, prompt explorer, and case gallery.

```bash
npm test
npm run lint
```

`npm test` builds the vinext/Cloudflare Worker application and runs artifact,
server-rendering, and Playwright interaction checks. Set `CHROMIUM_BIN` when a
`chromium` executable is not available on `PATH`.

## Main Artifacts

| Path | Purpose |
| --- | --- |
| `app/rubric/office-subquestion-results-v3.json` | Combined 54-case result artifact |
| `app/rubric/rubric-workbench.tsx` | Human/model comparison interface |
| `scripts/score_office_subquestion_v3.py` | Word/Excel rendering and scoring pipeline |
| `scripts/score_subquestion_ppt_v2.py` | 30-slide direct-subquestion scorer |
| `experiments/office-subquestion-v3/` | Manifest and resumable raw evidence |
| `public/benchmark-units-v3/` | Rendered Word and Excel units |
| `public/benchmark-slides-v2/` | Rendered PPT units |
| `tests/office-subquestion-v3.test.mjs` | V3 evidence-contract checks |

The Python scorers use `openai`, `pydantic`, and, for Office V3, `Pillow` plus
OfficeCLI. The checked-in result and progress artifacts can be inspected
without source documents or a model endpoint.

For a new run, `repo://` entries resolve against the OfficeCLI repository.
Place separately licensed source files under `source-documents/docx/` and
`source-documents/xlsx/`, or pass `--corpus-root`; these files are intentionally
ignored by Git. The PPT scorers read their source tree from
`OFFICE_REWARD_PPT_SOURCE_ROOT`. Published manifests retain source-document and
render hashes while using portable `repo://` and `corpus://` identifiers rather
than machine-specific paths.

## Experiment History

- V1: 12 PPT slides, criterion-level scoring, 120 scores and 60 abstentions.
- V2: 30 PPT slides, four direct subquestions per parent rubric block.
- V3: V2 plus 12 Word pages and 12 Excel sheets with format-specific rubrics.

Design decisions and implementation plans are under `docs/superpowers/`.
