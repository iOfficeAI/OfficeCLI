# Office Subquestion Scoring V3 Design

**Date:** 2026-08-11

**Status:** Approved by direct user instruction

## Goal

Extend the direct-subquestion experiment beyond PowerPoint with real Word and
Excel document units and format-specific parent blocks.

## Scope

V3 contains 54 cases: 30 reused PPT slides, 12 Word page-1 units, and 12 Excel
first-visible-sheet units. Word combines OfficeCLI feature documents and real
reports. Excel combines real operational workbooks and chart, formatting, and
pivot-table examples.

## Rubrics

Each format owns five Aesthetics and five Communication parent blocks. Every
parent contains four direct subquestions and a visible small score equal to the
mean of its children. PPT uses the V2 rubric unchanged. Word evaluates page
composition, typography, spacing, media integration, technical integrity,
purpose, progression, navigation, density, and audience fit. Excel evaluates
used-range composition, sizing and formats, input/output hierarchy, charts and
conditional formatting, scanability, workflow, labels, sheet context,
explainability, and decision relevance.

Content Accuracy remains `N/A` without an authoritative reference; coherence
is never substituted. Word and Excel have no imported human grades, so their
human fields remain blank and the page makes no alignment claim for them.

## Execution

OfficeCLI renders deterministic local PNGs and records source hashes and
provenance. GPT-5.5 performs two structured multimodal calls for each new case.
The 30 PPT results are reused exactly, producing 48 new calls and 108 total
logical calls represented in V3.

## Interface And Verification

Add format tabs for PPT, Word, and Excel while preserving the four-region
subquestion comparison. Verify 54 cases, 2,160 direct scores, 1,080 abstentions,
five parent scores, 20 rows per selected dimension, local human persistence,
desktop/mobile layout, and the public Cloudflare route.
