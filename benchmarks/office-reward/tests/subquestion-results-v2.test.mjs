import assert from "node:assert/strict";
import { access, readFile, stat } from "node:fs/promises";
import test from "node:test";

const artifactUrl = new URL(
  "../app/rubric/subquestion-results-v2.json",
  import.meta.url,
);

test("contains direct model scores for 30 slides and 40 subquestions each", async () => {
  const artifact = JSON.parse(await readFile(artifactUrl, "utf8"));
  assert.equal(artifact.experiment.id, "fine-grained-ppt-v2");
  assert.equal(artifact.experiment.case_count, 30);
  assert.equal(artifact.experiment.model_call_count, 60);
  assert.equal(artifact.experiment.subquestions_per_criterion, 4);
  assert.equal(artifact.experiment.scoring_mode, "two_dimension_subquestion_calls_per_slide");
  assert.match(artifact.experiment.rubric_sha256, /^[0-9a-f]{64}$/);
  assert.equal(artifact.summary.scored_subquestion_count, 1_200);
  assert.equal(artifact.summary.abstention_count, 600);
  assert.ok(artifact.summary.total_tokens > 0);
  assert.deepEqual(artifact.summary.human_grade_distribution, {
    0: 9,
    1: 9,
    2: 9,
    3: 3,
  });

  assert.equal(artifact.cases.length, 30);
  assert.equal(new Set(artifact.cases.map((item) => item.slide_id)).size, 30);
  for (const slideId of [
    "slide_0437", "slide_0459", "slide_0479", "slide_0002",
    "slide_0005", "slide_0014", "slide_0010", "slide_0026",
    "slide_0032", "slide_0387", "slide_0452", "slide_0025",
  ]) {
    assert.ok(artifact.cases.some((item) => item.slide_id === slideId));
  }

  for (const item of artifact.cases) {
    assert.match(item.image_sha256, /^[0-9a-f]{64}$/);
    assert.match(item.image, /^\/benchmark-slides-v2\/.+\.png$/);
    const image = new URL(`../public${item.image}`, import.meta.url);
    await access(image);
    assert.ok((await stat(image)).size > 10_000);
    assert.equal(item.results.length, 60);
    assert.equal(new Set(item.results.map((row) => row.subquestion_id)).size, 60);

    const scored = item.results.filter((row) => row.status === "scored");
    const abstained = item.results.filter(
      (row) => row.status === "not_assessable",
    );
    assert.equal(scored.length, 40);
    assert.equal(abstained.length, 20);
    for (const row of scored) {
      assert.equal(row.source, "model");
      assert.ok(Number.isInteger(row.score_1_5));
      assert.ok(row.score_1_5 >= 1 && row.score_1_5 <= 5);
      assert.ok(row.confidence_0_1 >= 0 && row.confidence_0_1 <= 1);
      assert.ok(row.evidence.trim().length >= 12);
      assert.ok(row.question.trim().length >= 8);
    }
    for (const row of abstained) {
      assert.equal(row.dimension, "content_accuracy");
      assert.equal(row.source, "abstention");
      assert.equal(row.score_1_5, null);
      assert.match(row.evidence, /No authoritative reference evidence/);
    }

    assert.equal(item.criterion_rollups.length, 15);
    for (const rollup of item.criterion_rollups) {
      const children = item.results.filter(
        (row) => row.criterion_id === rollup.criterion_id,
      );
      assert.equal(children.length, 4);
      if (rollup.dimension === "content_accuracy") {
        assert.equal(rollup.score_1_5, null);
      } else {
        const expected =
          Math.round(
            (children.reduce((sum, row) => sum + row.score_1_5, 0) / 4) *
              10_000,
          ) / 10_000;
        assert.equal(rollup.score_1_5, expected);
        assert.equal(rollup.source, "transparent_mean_of_direct_subquestions");
      }
    }
  }
});
