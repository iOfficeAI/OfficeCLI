"use client";

import Link from "next/link";
import { Fragment, useEffect, useMemo, useState } from "react";
import {
  dimensionLabels,
  dimensionOrder,
  formatLabels,
  formatOrder,
  officeSubquestionResults,
  type CriterionRollup,
  type DimensionKey,
  type OfficeFormat,
  type SubquestionResult,
} from "./office-subquestion-types";

type HumanScores = Record<string, number>;
type HumanNotes = Record<string, string>;

const STORAGE_KEY = "office-subquestion-human-scores-v3";

function storageKey(slideId: string, subquestionId: string) {
  return `${slideId}:${subquestionId}`;
}

function loadStoredState(): {
  humanScores: HumanScores;
  humanNotes: HumanNotes;
} {
  try {
    const stored = window.localStorage.getItem(STORAGE_KEY);
    if (!stored) return { humanScores: {}, humanNotes: {} };
    const parsed = JSON.parse(stored) as {
      humanScores?: HumanScores;
      humanNotes?: HumanNotes;
    };
    return {
      humanScores: parsed.humanScores ?? {},
      humanNotes: parsed.humanNotes ?? {},
    };
  } catch {
    return { humanScores: {}, humanNotes: {} };
  }
}

function scoreTone(score: number | null) {
  if (score === null) return "na";
  if (score >= 4) return "high";
  if (score === 3) return "mid";
  return "low";
}

function resultFor(
  results: SubquestionResult[],
  subquestionId: string,
): SubquestionResult {
  const result = results.find((item) => item.subquestion_id === subquestionId);
  if (!result) throw new Error(`Missing subquestion result: ${subquestionId}`);
  return result;
}

function rollupFor(
  rollups: CriterionRollup[],
  criterionId: string,
): CriterionRollup {
  const result = rollups.find((item) => item.criterion_id === criterionId);
  if (!result) throw new Error(`Missing criterion rollup: ${criterionId}`);
  return result;
}

export default function RubricWorkbench() {
  const [format, setFormat] = useState<OfficeFormat>("pptx");
  const [caseId, setCaseId] = useState(officeSubquestionResults.cases[0].case_uid);
  const [dimension, setDimension] = useState<DimensionKey>("aesthetics");
  const [hydrated, setHydrated] = useState(false);
  const [humanScores, setHumanScores] = useState<HumanScores>({});
  const [humanNotes, setHumanNotes] = useState<HumanNotes>({});

  const formatCases = useMemo(
    () => officeSubquestionResults.cases.filter((item) => item.format === format),
    [format],
  );
  const selectedIndex = Math.max(
    0,
    formatCases.findIndex((item) => item.case_uid === caseId),
  );
  const selectedCase = formatCases[selectedIndex];
  const criteria = officeSubquestionResults.rubric_by_format[selectedCase.format][dimension];
  const visibleSubquestions = useMemo(
    () =>
      criteria.flatMap((parent) =>
        parent.subquestions.map((subquestion, childIndex) => ({
          ...subquestion,
          criterionId: parent.id,
          criterionLabel: parent.label,
          childIndex,
        })),
      ),
    [criteria],
  );
  const visibleRollups = selectedCase.criterion_rollups.filter(
    (item) => item.dimension === dimension,
  );
  const completedCount = selectedCase.results.filter(
    (result) =>
      humanScores[
        storageKey(selectedCase.case_uid, result.subquestion_id)
      ] !== undefined,
  ).length;

  const meanAiScore = useMemo(() => {
    const values = selectedCase.results
      .filter((item) => item.dimension === dimension && item.score_1_5 !== null)
      .map((item) => Number(item.score_1_5));
    return values.length
      ? (values.reduce((sum, value) => sum + value, 0) / values.length).toFixed(2)
      : "N/A";
  }, [dimension, selectedCase]);

  useEffect(() => {
    const timer = window.setTimeout(() => {
      const stored = loadStoredState();
      setHumanScores(stored.humanScores);
      setHumanNotes(stored.humanNotes);
      setHydrated(true);
    }, 0);
    return () => window.clearTimeout(timer);
  }, []);

  useEffect(() => {
    if (!hydrated) return;
    window.localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({ humanScores, humanNotes }),
    );
  }, [humanNotes, humanScores, hydrated]);

  function selectRelativeCase(offset: number) {
    const nextIndex =
      (selectedIndex + offset + formatCases.length) % formatCases.length;
    setCaseId(formatCases[nextIndex].case_uid);
  }

  function selectFormat(nextFormat: OfficeFormat) {
    setFormat(nextFormat);
    const firstCase = officeSubquestionResults.cases.find(
      (item) => item.format === nextFormat,
    );
    if (firstCase) setCaseId(firstCase.case_uid);
    setDimension("aesthetics");
  }

  function updateScore(subquestionId: string, score: number) {
    const key = storageKey(selectedCase.case_uid, subquestionId);
    setHumanScores((current) => {
      if (current[key] === score) {
        const next = { ...current };
        delete next[key];
        return next;
      }
      return { ...current, [key]: score };
    });
  }

  function updateNote(subquestionId: string, note: string) {
    const key = storageKey(selectedCase.case_uid, subquestionId);
    setHumanNotes((current) => ({ ...current, [key]: note }));
  }

  function resetCurrentCase() {
    const prefix = `${selectedCase.case_uid}:`;
    setHumanScores((current) =>
      Object.fromEntries(
        Object.entries(current).filter(([key]) => !key.startsWith(prefix)),
      ),
    );
    setHumanNotes((current) =>
      Object.fromEntries(
        Object.entries(current).filter(([key]) => !key.startsWith(prefix)),
      ),
    );
  }

  function exportCurrentCase() {
    const rows = selectedCase.results.map((result) => {
      const key = storageKey(selectedCase.case_uid, result.subquestion_id);
      return {
        criterion_id: result.criterion_id,
        criterion: result.criterion_label,
        subquestion_id: result.subquestion_id,
        subquestion: result.subquestion_label,
        dimension: result.dimension,
        model_result: result,
        human_score_1_5: humanScores[key] ?? null,
        human_note: humanNotes[key] ?? "",
      };
    });
    const blob = new Blob(
      [
        JSON.stringify(
          {
            experiment: officeSubquestionResults.experiment,
            case: {
              case_uid: selectedCase.case_uid,
              format: selectedCase.format,
              slide_id: selectedCase.slide_id,
              case_id: selectedCase.case_id,
              image_sha256: selectedCase.image_sha256,
            },
            exported_at: new Date().toISOString(),
            criterion_rollups: selectedCase.criterion_rollups,
            rows,
          },
          null,
          2,
        ),
      ],
      { type: "application/json" },
    );
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${selectedCase.case_uid}-subquestion-human-scores-v3.json`;
    link.click();
    URL.revokeObjectURL(url);
  }

  return (
    <main className="fg-page">
      <header className="fg-topbar">
        <Link className="fg-brand" href="/rubric" aria-label="Office 细粒度评分实验">
          <span>V3</span>
          <strong>Office 细粒度评分实验</strong>
        </Link>
        <div className="fg-topbar-actions">
          <Link href="/">项目记录</Link>
          <span>1–5 DIRECT SUBQUESTION SCORES</span>
        </div>
      </header>

      <section className="fg-summary">
        <div className="fg-summary-copy">
          <p className="eyebrow">V3 / MULTI-FORMAT SUBQUESTION DIAGNOSTIC</p>
          <h1>PPT、Word、Excel，一套证据链</h1>
          <p>
            54 个真实 Office 单元：30 张 PPT、12 页 Word、12 张 Excel 工作表。
            每种格式使用自己的 5 个板块，每个板块拆成 4 个直接子问题。
          </p>
        </div>
        <dl className="fg-summary-stats">
          <div>
            <dt>直接子问题分</dt>
            <dd>{officeSubquestionResults.summary.scored_subquestion_count.toLocaleString("en-US")}</dd>
          </div>
          <div>
            <dt>模型调用</dt>
            <dd>{officeSubquestionResults.experiment.represented_model_call_count}</dd>
          </div>
          <div>
            <dt>PPT Spearman</dt>
            <dd>{officeSubquestionResults.summary.ppt_human_ai_aesthetics_spearman.toFixed(3)}</dd>
          </div>
          <div>
            <dt>PPT / Word / Excel</dt>
            <dd className="fg-distribution">30 / 12 / 12</dd>
          </div>
        </dl>
      </section>

      <section className="fg-diagnostic-band">
        <strong>真实诊断</strong>
        <span>
          PPT 有历史人工标签；Word 与 Excel 尚无人工总分，页面只展示真实模型分并等待人工子问题标注。
        </span>
        <span>Content Accuracy：无权威参考，1,080 个子问题全部 N/A</span>
      </section>

      <section className="fg-controls" aria-label="实验案例控制">
        <div className="fg-format-tabs" role="tablist" aria-label="Office 格式">
          {formatOrder.map((item) => (
            <button
              type="button"
              role="tab"
              aria-selected={format === item}
              className={format === item ? "active" : ""}
              onClick={() => selectFormat(item)}
              key={item}
            >
              {formatLabels[item]} · {officeSubquestionResults.summary.format_counts[item]}
            </button>
          ))}
        </div>
        <div className="fg-case-nav">
          <button type="button" aria-label="上一张" title="上一张" onClick={() => selectRelativeCase(-1)}>←</button>
          <label>
            <span>真实案例</span>
            <select value={selectedCase.case_uid} onChange={(event) => setCaseId(event.target.value)}>
              {formatCases.map((item, index) => (
                <option value={item.case_uid} key={item.case_uid}>
                  {String(index + 1).padStart(2, "0")} · {item.title}
                </option>
              ))}
            </select>
          </label>
          <button type="button" aria-label="下一张" title="下一张" onClick={() => selectRelativeCase(1)}>→</button>
        </div>

        <div className="fg-dimension-tabs" role="tablist" aria-label="评分维度">
          {dimensionOrder.map((item) => (
            <button
              type="button"
              role="tab"
              aria-selected={dimension === item}
              className={dimension === item ? "active" : ""}
              onClick={() => setDimension(item)}
              key={item}
            >
              {dimensionLabels[item].label}
            </button>
          ))}
        </div>

        <div className="fg-annotation-actions">
          <span>{completedCount} / 60 已填写</span>
          <button type="button" onClick={resetCurrentCase}>清空</button>
          <button type="button" onClick={exportCurrentCase}>导出 JSON</button>
        </div>
      </section>

      <section className="fg-case-strip">
        <div><span>{formatLabels[selectedCase.format]} CASE {selectedIndex + 1} / {formatCases.length}</span><strong>{selectedCase.case_uid}</strong></div>
        <div><span>来源</span><strong>{selectedCase.title}</strong></div>
        <div><span>评分单元</span><strong>{selectedCase.unit_type} · {selectedCase.unit_name}</strong></div>
        <div><span>历史人工美观分</span><strong>{selectedCase.human_aesthetics_grade_0_3 === null ? "未标注" : `${selectedCase.human_aesthetics_grade_0_3} / 3`}</strong></div>
        <div><span>当前维度均值</span><strong>{meanAiScore}{meanAiScore === "N/A" ? "" : " / 5"}</strong></div>
      </section>

      <section className="fg-rollup-strip" aria-label="板块小分">
        <div className="fg-rollup-label">
          <span>板块小分</span>
          <small>4 个直接子问题均值</small>
        </div>
        {visibleRollups.map((rollup) => (
          <div className={scoreTone(rollup.score_1_5)} key={rollup.criterion_id}>
            <span>{rollup.criterion_label}</span>
            <strong>{rollup.score_1_5 === null ? "N/A" : rollup.score_1_5.toFixed(2)}</strong>
          </div>
        ))}
      </section>

      <section className="fg-review-section">
        <div
          className="fg-review-grid"
          role="table"
          aria-label="Office 子问题评分对照表"
          style={{
            gridTemplateRows: `48px repeat(${visibleSubquestions.length}, minmax(214px, auto))`,
          }}
        >
          <div className="fg-column-head rule" role="columnheader">1. 评分细则 / 子问题</div>
          <div className="fg-column-head slide" role="columnheader">2. 真实 Office 单元</div>
          <div className="fg-column-head ai" role="columnheader">3. GPT-5.5 直接子问题分</div>
          <div className="fg-column-head human" role="columnheader">4. 人工子问题分</div>

          <aside className={`fg-slide-panel format-${selectedCase.format}`} style={{ gridRow: `2 / span ${visibleSubquestions.length}` }}>
            <div className="fg-slide-sticky">
              <img src={selectedCase.image} width={1280} height={720} alt={`${selectedCase.case_uid} rendered Office unit`} />
              <div className="fg-slide-caption">
                <p>{selectedCase.case_id.replaceAll("_", " ")}</p>
                <span>
                  {selectedCase.human_aesthetics_grade_0_3 === null
                    ? `${formatLabels[selectedCase.format]} · human label pending`
                    : `Human Aesthetics ${selectedCase.human_aesthetics_grade_0_3}/3`}
                </span>
              </div>
              <blockquote>{selectedCase.human_reason || "该案例没有附加人工缺陷说明。"}</blockquote>
            </div>
          </aside>

          {visibleSubquestions.map((subquestion, index) => {
            const result = resultFor(selectedCase.results, subquestion.id);
            const rollup = rollupFor(selectedCase.criterion_rollups, subquestion.criterionId);
            const key = storageKey(selectedCase.case_uid, subquestion.id);
            const humanScore = humanScores[key];
            const row = index + 2;
            const startsParent = subquestion.childIndex === 0;
            return (
              <Fragment key={subquestion.id}>
                <article
                  className={`fg-rule-cell ${startsParent ? "parent-start" : ""}`}
                  style={{ gridRow: row }}
                  role="cell"
                  data-subquestion-id={subquestion.id}
                >
                  <div className="fg-criterion-title">
                    <span>{String(index + 1).padStart(2, "0")}</span>
                    <div>
                      <small>{subquestion.criterionLabel} · {subquestion.childIndex + 1}/4</small>
                      <h2>{subquestion.label}</h2>
                    </div>
                  </div>
                  <p>{subquestion.question}</p>
                  <div className="fg-parent-score">
                    <span>板块小分</span>
                    <strong>{rollup.score_1_5 === null ? "N/A" : rollup.score_1_5.toFixed(2)}</strong>
                    <small>4 个直接子问题均值</small>
                  </div>
                  <div className="fg-subquestion-anchors">
                    <span><b>5</b> 有明确优秀证据，无实质缺陷</span>
                    <span><b>1</b> 严重失败、不可读或意图不可恢复</span>
                  </div>
                </article>

                <article
                  className={`fg-ai-cell ${scoreTone(result.score_1_5)} ${startsParent ? "parent-start" : ""}`}
                  style={{ gridRow: row }}
                  role="cell"
                >
                  <div className="fg-ai-score">
                    {result.score_1_5 === null ? <strong>N/A</strong> : <><strong>{result.score_1_5}</strong><span>/ 5</span></>}
                    <small>{result.status === "scored" ? `直接模型分 · 置信度 ${Math.round(result.confidence_0_1 * 100)}%` : "未提供参考证据"}</small>
                  </div>
                  <p>{result.evidence}</p>
                  {result.defects.length ? <ul>{result.defects.map((defect) => <li key={defect}>{defect}</li>)}</ul> : null}
                </article>

                <article className={`fg-human-cell ${startsParent ? "parent-start" : ""}`} style={{ gridRow: row }} role="cell">
                  <div className="fg-human-buttons" role="group" aria-label={`${subquestion.label}人工评分`}>
                    {[1, 2, 3, 4, 5].map((score) => (
                      <button type="button" aria-pressed={humanScore === score} onClick={() => updateScore(subquestion.id, score)} key={score}>{score}</button>
                    ))}
                  </div>
                  <textarea
                    value={humanNotes[key] ?? ""}
                    onChange={(event) => updateNote(subquestion.id, event.target.value)}
                    rows={3}
                    aria-label={`${subquestion.label}人工备注`}
                    placeholder="人工证据或备注（可选）"
                  />
                  <span>{humanScore ? `已选 ${humanScore} 分` : "未填写"}</span>
                </article>
              </Fragment>
            );
          })}
        </div>
      </section>

      <footer className="fg-footer">
        <span>{officeSubquestionResults.experiment.model} · {officeSubquestionResults.experiment.reasoning_effort} reasoning · {officeSubquestionResults.summary.represented_total_tokens.toLocaleString("en-US")} represented tokens</span>
        <span>Rubric {officeSubquestionResults.experiment.rubric_sha256.slice(0, 12)} · {officeSubquestionResults.experiment.generated_at.slice(0, 10)}</span>
      </footer>
    </main>
  );
}
