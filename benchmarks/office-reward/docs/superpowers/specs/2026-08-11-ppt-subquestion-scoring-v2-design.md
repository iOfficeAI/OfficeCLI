# PPT Subquestion Scoring V2 Design

**Date:** 2026-08-11

**Status:** Approved by direct user instruction

## Goal

Expand the real PPT scoring diagnostic from 12 to 30 slides and replace each
criterion-level judgment with four independently scored subquestions.

## Sample

The source annotations contain only three valid Human Grade 3 slides. V2 keeps
all 12 V1 slides, then deterministically adds six unique-case slides for each
of grades 0, 1, and 2. The final distribution is `0:9, 1:9, 2:9, 3:3`.
The imbalance is reported and never described as a balanced four-grade set.

## Scoring

Aesthetics and Communication Effectiveness each retain five parent criteria.
Every parent contains four concrete subquestions. One multimodal request per
dimension returns 20 direct 1-5 subquestion scores, each with confidence,
visible evidence, and at most two defects. Parent criterion scores are
transparent arithmetic means of their four direct subquestion scores.

Content Accuracy also has four subquestions per parent, but all 20 abstain
without authoritative reference evidence. Internal coherence is not scored as
accuracy.

## Interface

The existing four visual regions remain. Each selected dimension now shows 20
rows: parent criterion plus one subquestion, sticky slide image, direct model
subquestion score, and blank human subquestion score. Human state uses a new V2
storage namespace and exports all direct model provenance.

## Verification

- Exactly 30 cases, 1,200 model-scored subquestions, and 600 abstentions.
- Exactly 60 logical model calls with resume metadata bound to model and rubric.
- Browser checks cover subquestion rows, case isolation, N/A accuracy rows,
  image loading, desktop columns, mobile overflow, and public Cloudflare access.
