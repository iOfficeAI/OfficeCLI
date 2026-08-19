# Real Fine-Grained PPT Scoring Design

**Date:** 2026-08-11

**Status:** Approved for implementation

## Goal

Replace derived AI subscores with reproducible model assessments for 12 real
benchmark slides while keeping human criterion scores blank and editable.

## Contract

The tracked dataset contains three slides at each existing human Aesthetics
grade from 0 through 3. Each slide receives one real multimodal evaluation for
five Aesthetics criteria and one for five Communication criteria. Every result
contains a 1-5 score, confidence, visible evidence, and defects.

The five Content Accuracy criteria are `not_assessable` because the benchmark
does not contain authoritative references. They must never receive inferred or
coherence-based accuracy scores. Saved results include model configuration,
rubric and image hashes, timestamp, and the criterion outputs.

## Interface

The `/rubric` route uses four visual regions: rubric criterion, current slide
image, real AI result, and blank human score plus note. The image remains
visible while five criteria are reviewed. Human entries persist locally and
export with immutable model provenance. Narrow screens stack without overflow.

## Integrity

- Remove the `raw dimension score / 20 + bias` path.
- Require 12 cases and 15 rows per case: ten scored and five abstained.
- Preserve model failures; never substitute fixture scores.
- Keep unrelated working-tree changes intact.
- Require contract, browser, and production-build verification.
