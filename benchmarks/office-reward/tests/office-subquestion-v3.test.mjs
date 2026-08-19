import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { access, readFile, stat } from "node:fs/promises";
import test from "node:test";

const artifactUrl = new URL(
  "../app/rubric/office-subquestion-results-v3.json",
  import.meta.url,
);
const manifestUrl = new URL(
  "../experiments/office-subquestion-v3/manifest.json",
  import.meta.url,
);
const progressUrl = new URL(
  "../experiments/office-subquestion-v3/progress.json",
  import.meta.url,
);

test("publishes portable and internally consistent source provenance", async () => {
  const manifestText = await readFile(manifestUrl, "utf8");
  const manifest = JSON.parse(manifestText);
  const progress = JSON.parse(await readFile(progressUrl, "utf8"));
  const artifact = JSON.parse(await readFile(artifactUrl, "utf8"));
  const artifactById = new Map(artifact.cases.map((item) => [item.case_uid, item]));

  assert.equal(manifest.length, 24);
  assert.equal(
    createHash("sha256").update(manifestText).digest("hex"),
    progress.metadata.manifest_sha256,
  );

  for (const row of manifest) {
    assert.match(row.source, /^(repo|corpus):\/\//);
    assert.doesNotMatch(row.source, /^(\/|[A-Za-z]:\\)/);
    assert.equal(
      row.source_document_sha256,
      artifactById.get(row.case_uid).source_document_sha256,
    );
    if (row.source.startsWith("repo://")) {
      const source = new URL(`../../../${row.source.slice(7)}`, import.meta.url);
      await access(source);
    }
  }
});

test("contains direct subquestion evidence for PPT, Word, and Excel", async () => {
  const artifact = JSON.parse(await readFile(artifactUrl, "utf8"));
  assert.equal(artifact.experiment.id, "office-subquestion-v3");
  assert.equal(artifact.experiment.case_count, 54);
  assert.equal(artifact.experiment.new_model_call_count, 48);
  assert.equal(artifact.experiment.represented_model_call_count, 108);
  assert.deepEqual(artifact.summary.format_counts, {
    pptx: 30,
    docx: 12,
    xlsx: 12,
  });
  assert.equal(artifact.summary.scored_subquestion_count, 2_160);
  assert.equal(artifact.summary.abstention_count, 1_080);
  assert.ok(artifact.summary.new_total_tokens > 0);

  assert.deepEqual(Object.keys(artifact.rubric_by_format), [
    "pptx",
    "docx",
    "xlsx",
  ]);
  for (const format of Object.keys(artifact.rubric_by_format)) {
    for (const dimension of [
      "aesthetics",
      "content_accuracy",
      "communication_effectiveness",
    ]) {
      assert.equal(artifact.rubric_by_format[format][dimension].length, 5);
      assert.ok(
        artifact.rubric_by_format[format][dimension].every(
          (criterion) => criterion.subquestions.length === 4,
        ),
      );
    }
  }

  assert.equal(artifact.cases.length, 54);
  assert.equal(new Set(artifact.cases.map((item) => item.case_uid)).size, 54);
  for (const item of artifact.cases) {
    assert.ok(["pptx", "docx", "xlsx"].includes(item.format));
    assert.equal(item.results.length, 60);
    assert.equal(item.criterion_rollups.length, 15);
    assert.equal(item.results.filter((row) => row.status === "scored").length, 40);
    assert.equal(
      item.results.filter((row) => row.status === "not_assessable").length,
      20,
    );
    assert.match(item.image_sha256, /^[0-9a-f]{64}$/);
    const image = new URL(`../public${item.image}`, import.meta.url);
    await access(image);
    assert.ok((await stat(image)).size > 10_000);

    if (item.format === "pptx") {
      assert.equal(item.evidence_source, "v2_reused");
      assert.ok(Number.isInteger(item.human_aesthetics_grade_0_3));
    } else {
      assert.equal(item.evidence_source, "v3_new_model_calls");
      assert.equal(item.human_aesthetics_grade_0_3, null);
      assert.match(item.source_document_sha256, /^[0-9a-f]{64}$/);
      assert.ok(item.source_name.endsWith(`.${item.format}`));
      assert.match(item.image, /^\/benchmark-units-v3\/.+\.png$/);
    }

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
      }
    }
  }
});
