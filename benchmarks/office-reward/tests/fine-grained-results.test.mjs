import assert from "node:assert/strict";
import { access, readFile, stat } from "node:fs/promises";
import test from "node:test";

const artifactUrl = new URL(
  "../app/rubric/fine-grained-results.json",
  import.meta.url,
);

test("contains real criterion-level model scores for 12 benchmark slides", async () => {
  const artifact = JSON.parse(await readFile(artifactUrl, "utf8"));

  assert.equal(artifact.experiment.id, "fine-grained-ppt-v1");
  assert.equal(artifact.experiment.case_count, 12);
  assert.equal(artifact.experiment.model_call_count, 24);
  assert.equal(artifact.experiment.scoring_mode, "two_dimension_calls_per_slide");
  assert.match(artifact.experiment.generated_at, /^2026-/);
  assert.ok(artifact.experiment.model.length > 0);
  assert.match(artifact.experiment.rubric_sha256, /^[0-9a-f]{64}$/);
  assert.equal(artifact.summary.scored_criterion_count, 120);
  assert.equal(artifact.summary.abstention_count, 60);
  assert.equal(artifact.summary.total_tokens, 69_712);
  assert.ok(artifact.summary.human_ai_aesthetics_spearman > 0.5);
  assert.ok(artifact.summary.human_ai_aesthetics_spearman < 0.6);
  assert.deepEqual(Object.keys(artifact.summary.mean_aesthetics_by_human_grade), [
    "0",
    "1",
    "2",
    "3",
  ]);

  assert.equal(artifact.cases.length, 12);
  assert.equal(new Set(artifact.cases.map((item) => item.slide_id)).size, 12);

  const gradeCounts = {};
  for (const item of artifact.cases) {
    gradeCounts[item.human_aesthetics_grade_0_3] =
      (gradeCounts[item.human_aesthetics_grade_0_3] ?? 0) + 1;
    assert.match(item.image_sha256, /^[0-9a-f]{64}$/);
    assert.match(item.image, /^\/benchmark-slides\/.+\.png$/);
    const image = new URL(`../public${item.image}`, import.meta.url);
    await access(image);
    assert.ok((await stat(image)).size > 10_000);

    assert.equal(item.results.length, 15);
    assert.equal(
      new Set(item.results.map((result) => result.criterion_id)).size,
      15,
    );
    const scored = item.results.filter((result) => result.status === "scored");
    const abstained = item.results.filter(
      (result) => result.status === "not_assessable",
    );
    assert.equal(scored.length, 10);
    assert.equal(abstained.length, 5);

    for (const result of scored) {
      assert.equal(result.source, "model");
      assert.ok(Number.isInteger(result.score_1_5));
      assert.ok(result.score_1_5 >= 1 && result.score_1_5 <= 5);
      assert.ok(result.confidence_0_1 >= 0 && result.confidence_0_1 <= 1);
      assert.ok(result.evidence.trim().length >= 20);
      assert.ok(Array.isArray(result.defects));
    }
    for (const result of abstained) {
      assert.equal(result.dimension, "content_accuracy");
      assert.equal(result.source, "abstention");
      assert.equal(result.score_1_5, null);
      assert.match(result.evidence, /No authoritative reference evidence/);
    }
  }

  assert.deepEqual(gradeCounts, { 0: 3, 1: 3, 2: 3, 3: 3 });
  assert.doesNotMatch(JSON.stringify(artifact), /derived|dimension_score\s*\/\s*20/i);
});
