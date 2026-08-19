"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { rewardCases, type OfficeFormat, type RewardCase } from "./case-data";

type Filter = "all" | OfficeFormat;

const filters: Array<{ key: Filter; label: string }> = [
  { key: "all", label: "全部" },
  { key: "pptx", label: "PPT" },
  { key: "docx", label: "Word" },
  { key: "xlsx", label: "Excel" },
];

const formatLabels: Record<OfficeFormat, string> = {
  pptx: "PowerPoint",
  docx: "Word",
  xlsx: "Excel",
};

const scoreRows = [
  ["aesthetics", "Aesthetics"],
  ["content_accuracy", "Content Accuracy"],
  ["communication_effectiveness", "Communication"],
] as const;

export default function CaseGallery() {
  const [filter, setFilter] = useState<Filter>("all");
  const [selected, setSelected] = useState<RewardCase | null>(null);
  const dialogRef = useRef<HTMLDialogElement>(null);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const visibleCases = useMemo(
    () =>
      filter === "all"
        ? rewardCases
        : rewardCases.filter((item) => item.format === filter),
    [filter],
  );

  useEffect(() => {
    if (!selected) return;
    const dialog = dialogRef.current;
    if (!dialog) return;
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    if (!dialog.open) dialog.showModal();
    return () => {
      if (dialog.open) dialog.close();
      document.body.style.overflow = previousOverflow;
      triggerRef.current?.focus();
    };
  }, [selected]);

  return (
    <section className="section cases-section" id="cases">
      <div className="section-heading case-heading">
        <div>
          <p className="eyebrow">CASE GALLERY</p>
          <h2>完整评分案例</h2>
        </div>
        <p>九个真实渲染案例，统一展示三维分数、reward、coverage 与证据。</p>
      </div>

      <div className="case-toolbar">
        <div className="case-filters" aria-label="案例格式筛选">
          {filters.map((item) => (
            <button
              type="button"
              className={filter === item.key ? "active" : ""}
              aria-pressed={filter === item.key}
              onClick={() => setFilter(item.key)}
              key={item.key}
            >
              {item.label}
            </button>
          ))}
        </div>
        <span className="case-count">{visibleCases.length} cases</span>
      </div>

      <div className="case-grid">
        {visibleCases.map((item) => (
          <article className={`case-card ${item.format}`} key={item.id}>
            <div className="case-card-open">
              <div className="case-image">
                <img
                  src={item.image}
                  width={1600}
                  height={1200}
                  alt={`${item.title} Office reward case`}
                  loading="lazy"
                  decoding="async"
                />
              </div>
              <div className="case-content">
                <div className="case-kicker">
                  <span className="case-format">{formatLabels[item.format]}</span>
                  <span className="case-status">{item.status}</span>
                </div>
                <h3>{item.title}</h3>
                <p>{item.subtitle}</p>
                <div className="case-scores">
                  {scoreRows.map(([key, label]) => (
                    <div className="case-score-row" key={key}>
                      <span>{label}</span>
                      <i>
                        <b style={{ width: `${item.scores[key]}%` }} />
                      </i>
                      <strong>{item.scores[key]}</strong>
                    </div>
                  ))}
                </div>
                <dl className="case-summary">
                  <div>
                    <dt>Reward</dt>
                    <dd>{item.reward_0_1.toFixed(4)}</dd>
                  </div>
                  <div>
                    <dt>Coverage</dt>
                    <dd>{Math.round(item.coverage_0_1 * 100)}%</dd>
                  </div>
                  <div>
                    <dt>Units</dt>
                    <dd>{item.unitCount}</dd>
                  </div>
                </dl>
              </div>
            </div>
            <button
              type="button"
              className="case-detail-button"
              aria-haspopup="dialog"
              onClick={(event) => {
                triggerRef.current = event.currentTarget;
                setSelected(item);
              }}
            >
              查看详情 <span aria-hidden="true">→</span>
            </button>
          </article>
        ))}
      </div>

      {selected ? (
        <dialog
          ref={dialogRef}
          className="case-dialog-shell"
          role="dialog"
          aria-modal="true"
          aria-labelledby="case-dialog-title"
          onCancel={() => setSelected(null)}
          onClose={() => setSelected(null)}
          onMouseDown={(event) => {
            if (event.target === event.currentTarget) {
              dialogRef.current?.close();
            }
          }}
        >
          <div
            className={`case-dialog ${selected.format}`}
          >
            <button
              type="button"
              className="case-dialog-close"
              aria-label="关闭案例详情"
              autoFocus
              onClick={() => dialogRef.current?.close()}
            >
              ×
            </button>
            <div className="case-dialog-media">
              <img
                src={selected.image}
                width={1600}
                height={1200}
                alt={`${selected.title} 详情`}
                decoding="async"
              />
            </div>
            <div className="case-dialog-body">
              <p className="eyebrow">{formatLabels[selected.format]} CASE</p>
              <h2 id="case-dialog-title">{selected.title}</h2>
              <p className="case-dialog-subtitle">{selected.subtitle}</p>

              <div className="dialog-score-grid">
                {scoreRows.map(([key, label]) => (
                  <div key={key}>
                    <span>{label}</span>
                    <strong>{selected.scores[key]}</strong>
                  </div>
                ))}
              </div>

              <div className="reward-equation">
                <span>Raw overall</span>
                <strong>{selected.overall_raw_score_100.toFixed(2)}</strong>
                <span>÷ 100</span>
                <strong>{selected.reward_0_1.toFixed(4)}</strong>
              </div>

              <dl className="dialog-metadata">
                <div>
                  <dt>Unit type</dt>
                  <dd>{selected.unitType}</dd>
                </div>
                <div>
                  <dt>Unit count</dt>
                  <dd>{selected.unitCount}</dd>
                </div>
                <div>
                  <dt>Coverage</dt>
                  <dd>{Math.round(selected.coverage_0_1 * 100)}%</dd>
                </div>
                <div>
                  <dt>Status</dt>
                  <dd>{selected.status}</dd>
                </div>
              </dl>

              <div className="dialog-evidence">
                <h3>Evidence</h3>
                <ul>
                  {selected.evidence.map((item) => (
                    <li key={item}>{item}</li>
                  ))}
                </ul>
              </div>

              <div className="dialog-issues">
                <span>Office issue summary</span>
                <p>{selected.issueSummary}</p>
              </div>
            </div>
          </div>
        </dialog>
      ) : null}
    </section>
  );
}
