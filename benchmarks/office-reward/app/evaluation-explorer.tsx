"use client";

import { useState } from "react";
import evaluationContent from "./evaluation-content.json";

type FormatKey = keyof typeof evaluationContent.prompts;
type DimensionKey =
  keyof (typeof evaluationContent.prompts)[FormatKey]["dimensions"];

const formatOrder: FormatKey[] = ["pptx", "docx", "xlsx"];
const dimensionOrder: DimensionKey[] = [
  "aesthetics",
  "content_accuracy",
  "communication_effectiveness",
];

const formatShortLabels: Record<FormatKey, string> = {
  pptx: "PPT",
  docx: "Word",
  xlsx: "Excel",
};

const dimensionShortLabels: Record<DimensionKey, string> = {
  aesthetics: "Aesthetics",
  content_accuracy: "Content Accuracy",
  communication_effectiveness: "Communication",
};

const weightRows = [
  {
    key: "aesthetics",
    label: "Aesthetics",
    percent: evaluationContent.weights.aesthetics * 100,
  },
  {
    key: "content_accuracy",
    label: "Content Accuracy",
    percent: evaluationContent.weights.content_accuracy * 100,
  },
  {
    key: "communication_effectiveness",
    label: "Communication",
    percent: evaluationContent.weights.communication_effectiveness * 100,
  },
] as const;

async function copyText(text: string) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.style.position = "fixed";
  textarea.style.opacity = "0";
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand("copy");
  textarea.remove();
}

export default function EvaluationExplorer() {
  const [format, setFormat] = useState<FormatKey>("pptx");
  const [dimension, setDimension] =
    useState<DimensionKey>("aesthetics");
  const [copyState, setCopyState] = useState<"idle" | "copied" | "failed">(
    "idle",
  );

  const selectedFormat = evaluationContent.prompts[format];
  const selectedPrompt = selectedFormat.dimensions[dimension];
  const multiDimensionInstructions =
    evaluationContent.multiDimensionInstructions;

  async function handleCopy() {
    try {
      await copyText(selectedPrompt.prompt);
      setCopyState("copied");
      window.setTimeout(() => setCopyState("idle"), 1600);
    } catch {
      setCopyState("failed");
      window.setTimeout(() => setCopyState("idle"), 1600);
    }
  }

  return (
    <>
      <section className="section mechanism-section" id="mechanism">
        <div className="section-heading">
          <div>
            <p className="eyebrow">EVALUATION MECHANISM</p>
            <h2>评测机制</h2>
          </div>
          <p>
            OfficeCLI 先生成可核验的文档单元，再做三维独立评分、文档聚合与 reward
            输出。
          </p>
        </div>

        <div className="mechanism-flow">
          {evaluationContent.stages.map((stage) => (
            <article className="mechanism-step" key={stage.number}>
              <span>{stage.number}</span>
              <div>
                <h3>{stage.title}</h3>
                <p>{stage.text}</p>
              </div>
            </article>
          ))}
        </div>

        <div className="mechanism-details">
          <div className="weight-panel">
            <div className="panel-heading">
              <span>Dimension weights</span>
              <strong>100%</strong>
            </div>
            {weightRows.map((item) => (
              <div className="weight-row" key={item.key}>
                <span>{item.label}</span>
                <i>
                  <b style={{ width: `${item.percent}%` }} />
                </i>
                <strong>{item.percent}%</strong>
              </div>
            ))}
          </div>

          <div className="formula-panel">
            <div className="panel-heading">
              <span>Aggregation formulas</span>
              <strong>4</strong>
            </div>
            <ol>
              {evaluationContent.formulas.map((formula) => (
                <li key={formula}>
                  <code>{formula}</code>
                </li>
              ))}
            </ol>
          </div>
        </div>

        <div className="status-definitions">
          {evaluationContent.statuses.map((status) => (
            <article key={status.name}>
              <span className={`evaluation-status ${status.name}`}>
                {status.name}
              </span>
              <p>{status.text}</p>
            </article>
          ))}
        </div>
      </section>

      <section className="section prompt-section" id="prompts">
        <div className="section-heading">
          <div>
            <p className="eyebrow">PROMPT EXPLORER</p>
            <h2>实际评测 Prompt</h2>
          </div>
          <p>
            直接展示当前评分代码生成的完整 rubric，覆盖 PPT、Word、Excel
            的三个独立维度。
          </p>
        </div>

        <div className="prompt-toolbar">
          <div
            className="prompt-tabs format-tabs"
            role="tablist"
            aria-label="选择 Office 格式"
          >
            {formatOrder.map((key) => (
              <button
                type="button"
                role="tab"
                aria-selected={format === key}
                className={format === key ? "active" : ""}
                onClick={() => {
                  setFormat(key);
                  setCopyState("idle");
                }}
                key={key}
              >
                {formatShortLabels[key]}
              </button>
            ))}
          </div>

          <div
            className="prompt-tabs dimension-tabs"
            role="tablist"
            aria-label="选择评分维度"
          >
            {dimensionOrder.map((key) => (
              <button
                type="button"
                role="tab"
                aria-selected={dimension === key}
                className={dimension === key ? "active" : ""}
                onClick={() => {
                  setDimension(key);
                  setCopyState("idle");
                }}
                key={key}
              >
                {dimensionShortLabels[key]}
              </button>
            ))}
          </div>
        </div>

        <div className="prompt-viewer">
          <div className="prompt-viewer-header">
            <div>
              <span>
                {selectedFormat.label} / {selectedPrompt.label}
              </span>
              <code>SHA256 {selectedPrompt.hash}</code>
            </div>
            <button type="button" onClick={handleCopy}>
              {copyState === "copied"
                ? "已复制"
                : copyState === "failed"
                  ? "复制失败"
                  : "复制 Prompt"}
            </button>
          </div>
          <pre
            role="tabpanel"
            aria-label={`${selectedFormat.label} ${selectedPrompt.label} prompt`}
          >
            <code>{selectedPrompt.prompt}</code>
          </pre>
        </div>

        <div className="prompt-supporting">
          <details>
            <summary>
              <span>Request wrapper</span>
              <small>模型请求上下文模板</small>
            </summary>
            <pre>
              <code>{evaluationContent.requestWrapper}</code>
            </pre>
          </details>
          <details>
            <summary>
              <span>Multi-dimension instructions</span>
              <small>一次响应内保持三维判断独立</small>
            </summary>
            <pre>
              <code>{multiDimensionInstructions}</code>
            </pre>
          </details>
        </div>
      </section>
    </>
  );
}
