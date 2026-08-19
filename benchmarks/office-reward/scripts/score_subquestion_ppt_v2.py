#!/usr/bin/env python3
"""Run direct subquestion-level PPT evaluation on a fixed 30-slide sample."""

from __future__ import annotations

import argparse
import json
import shutil
import threading
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openai import OpenAI
from pydantic import BaseModel, ConfigDict, Field, model_validator

from score_fine_grained_ppt import (
    ABSTENTION_REASON,
    REPO_ROOT,
    SOURCE_ROOT,
    atomic_json,
    canonical_json,
    image_data_url,
    sha256_bytes,
    spearman,
)


DEFAULT_ANNOTATIONS = SOURCE_ROOT / "data/slide_annotations.jsonl"
DEFAULT_V1_MANIFEST = SOURCE_ROOT / "configs/scoring_case_manifest.jsonl"
DEFAULT_OUTPUT = REPO_ROOT / "app/rubric/subquestion-results-v2.json"
DEFAULT_ASSET_DIR = REPO_ROOT / "public/benchmark-slides-v2"
DEFAULT_RUN_DIR = REPO_ROOT / "experiments/fine-grained-ppt-v2"
RUBRIC_VERSION = "ppt-subquestion-rubric-v2"


def criterion(
    criterion_id: str,
    label: str,
    subquestions: list[tuple[str, str, str]],
) -> dict[str, Any]:
    return {
        "id": criterion_id,
        "label": label,
        "subquestions": [
            {
                "id": f"{criterion_id}-{suffix}",
                "label": sublabel,
                "question": question,
            }
            for suffix, sublabel, question in subquestions
        ],
    }


RUBRIC: dict[str, list[dict[str, Any]]] = {
    "aesthetics": [
        criterion(
            "ppt-layout",
            "布局与构图",
            [
                ("balance", "视觉平衡", "页面左右、上下的视觉重量是否平衡，是否存在明显头重脚轻或偏置？"),
                ("alignment", "对齐关系", "标题、文本、图表和卡片是否沿稳定参考线对齐？"),
                ("margins", "边距与画布", "安全边距是否一致，16:9 画布是否被有意且充分地使用？"),
                ("grouping", "分组结构", "相关元素是否形成清晰组块，不相关内容是否有足够分隔？"),
            ],
        ),
        criterion(
            "ppt-typography",
            "字体与层级",
            [
                ("readability", "演示可读性", "正文、标签和注释在演示距离下是否清楚可读？"),
                ("hierarchy", "标题层级", "标题、副标题、正文和注释是否形成明确稳定的层级？"),
                ("consistency", "字体一致性", "字体、字号、字重和样式是否在同类元素间保持一致？"),
                ("contrast", "行长与对比度", "文字行长、背景对比和段落密度是否支持快速阅读？"),
            ],
        ),
        criterion(
            "ppt-graphics",
            "图表与图形",
            [
                ("legibility", "图形可辨性", "图片、图表、表格和图标是否清晰且关键细节可辨？"),
                ("labels", "标注完整性", "图例、坐标轴、数据标签、箭头和连接线是否完整准确地呈现？"),
                ("relevance", "内容相关性", "每个图形是否直接服务于当前页面的主要信息？"),
                ("integration", "风格与整合", "图形风格、尺寸、颜色和位置是否与页面其余部分协调？"),
            ],
        ),
        criterion(
            "ppt-technical",
            "技术完整性",
            [
                ("overlap", "元素重叠", "是否存在会遮挡文字或图形、影响理解的非预期重叠？"),
                ("clipping", "裁切与截断", "文字、图片、图表或表格是否被裁切、截断或显示不全？"),
                ("overflow", "边界溢出", "元素是否越出画布、容器或安全边距，造成不可读区域？"),
                ("artifacts", "渲染与连接故障", "是否存在断裂连接线、缺失资源、乱码或明显渲染伪影？"),
            ],
        ),
        criterion(
            "ppt-economy",
            "视觉经济性",
            [
                ("density", "信息密度", "信息量是否适中，既不文档式拥挤也不过度稀疏？"),
                ("whitespace", "留白质量", "留白是否帮助分组和聚焦，而不是形成无理由的大块空洞？"),
                ("redundancy", "视觉冗余", "是否存在重复图表、重复文字或无贡献的重复组件？"),
                ("decoration", "装饰克制", "装饰元素是否克制，不与核心信息争夺注意力？"),
            ],
        ),
    ],
    "content_accuracy": [
        criterion(
            "ppt-central-claim",
            "中心结论",
            [
                ("identity", "结论识别", "页面中心事实性结论能否被明确识别？"),
                ("direction", "方向正确性", "中心结论的提升、下降或比较方向是否与参考证据一致？"),
                ("scope", "结论范围", "中心结论的适用范围和限定条件是否准确？"),
                ("support", "直接支持", "中心结论是否得到权威参考证据直接支持？"),
            ],
        ),
        criterion(
            "ppt-numbers",
            "数字与计算",
            [
                ("values", "数值一致性", "可见数值和百分比是否与参考证据一致？"),
                ("units", "单位与精度", "单位、符号、小数精度和舍入是否准确一致？"),
                ("calculations", "计算关系", "公式、差值、合计和派生计算是否正确？"),
                ("chart-table", "图表表格一致性", "图表、表格与正文引用的数值是否彼此一致？"),
            ],
        ),
        criterion(
            "ppt-labels",
            "标签与来源",
            [
                ("entities", "实体命名", "方法、模型、数据集、组织和专有名称是否准确？"),
                ("chart-labels", "图表标签", "坐标轴、图例、系列、类别和表头是否准确完整？"),
                ("attribution", "来源归因", "外部结果、图片和观点是否归因到正确来源？"),
                ("citations", "引用可定位性", "引用和来源说明是否足以定位支撑材料？"),
            ],
        ),
        criterion(
            "ppt-scope",
            "范围与因果",
            [
                ("baseline", "比较基线", "比较对象、基线和实验条件是否表述准确？"),
                ("causality", "因果强度", "相关性、贡献和因果关系是否被正确区分？"),
                ("superlative", "最高级声明", "最佳、首次、全部等强声明是否有充分证据？"),
                ("generalization", "泛化边界", "页面是否避免把局部结果不当地推广到更大范围？"),
            ],
        ),
        criterion(
            "ppt-reference-coverage",
            "证据覆盖",
            [
                ("central", "核心覆盖", "参考材料是否覆盖页面中心结论？"),
                ("numbers", "数字覆盖", "参考材料是否覆盖所有重要可见数字？"),
                ("supporting", "支撑声明覆盖", "重要支撑声明是否能逐项核验？"),
                ("residual", "剩余不可核验项", "页面是否清楚区分仍不可核验的内容？"),
            ],
        ),
    ],
    "communication_effectiveness": [
        criterion(
            "ppt-takeaway",
            "核心信息清晰度",
            [
                ("recoverability", "一句话复述", "仅看页面能否用一句话准确复述核心信息？"),
                ("uniqueness", "信息唯一性", "页面是否存在唯一稳定主线，而非多个竞争性结论？"),
                ("title", "标题传意", "标题是否直接表达或准确引导核心信息？"),
                ("speed", "理解速度", "目标受众能否在数秒内抓住页面要点？"),
            ],
        ),
        criterion(
            "ppt-alignment",
            "标题与证据一致性",
            [
                ("title-claim", "标题与结论", "标题和页面实际结论是否同向？"),
                ("emphasis", "强调与优先级", "视觉强调是否落在最重要的信息上？"),
                ("evidence", "证据相关性", "主要证据是否直接支持标题和核心结论？"),
                ("conclusion", "结论一致性", "底部总结或行动项是否与前文证据一致？"),
            ],
        ),
        criterion(
            "ppt-reading-path",
            "阅读路径",
            [
                ("entry", "第一落点", "页面是否有明确且合理的第一视觉落点？"),
                ("sequence", "阅读顺序", "第二、第三信息点是否按自然顺序推进？"),
                ("transitions", "关系与过渡", "箭头、编号、分组和位置是否清楚表达内容关系？"),
                ("landing", "最终落点", "阅读路径是否落到明确结论、结果或行动上？"),
            ],
        ),
        criterion(
            "ppt-comprehension",
            "受众理解成本",
            [
                ("terms", "术语解释", "专业术语和缩写是否得到足够解释？"),
                ("legends", "图例自解释性", "图例、符号、颜色和标签是否无需猜测即可理解？"),
                ("burden", "阅读负担", "字号、密度和句子长度是否避免反复细读？"),
                ("context", "受众上下文", "页面是否提供目标受众理解结论所需的背景？"),
            ],
        ),
        criterion(
            "ppt-information-economy",
            "信息取舍",
            [
                ("relevance", "内容必要性", "每个主要内容块是否直接服务于核心信息？"),
                ("duplication", "重复控制", "是否避免表格、图表和文字重复表达同一证据？"),
                ("detail", "细节层级", "次要细节是否被压缩或移出主页面，而非淹没主线？"),
                ("ornament", "装饰竞争", "装饰是否避免抢占本应属于证据和结论的注意力？"),
            ],
        ),
    ],
}


class SubquestionAssessment(BaseModel):
    model_config = ConfigDict(extra="forbid")

    subquestion_id: str
    score_1_5: int = Field(ge=1, le=5)
    confidence_0_1: float = Field(ge=0, le=1)
    evidence: str = Field(min_length=12, max_length=260)
    defects: list[str] = Field(default_factory=list, max_length=2)


class DimensionAssessment(BaseModel):
    model_config = ConfigDict(extra="forbid")

    results: list[SubquestionAssessment] = Field(min_length=20, max_length=20)

    @model_validator(mode="after")
    def unique_subquestions(self) -> "DimensionAssessment":
        identifiers = [item.subquestion_id for item in self.results]
        if len(set(identifiers)) != 20:
            raise ValueError("subquestion IDs must contain 20 unique values")
        return self


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    return [
        json.loads(line)
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]


def normalize_v1_row(row: dict[str, Any]) -> dict[str, Any]:
    metadata = row["metadata"]
    return {
        "slide_id": row["slide_id"],
        "image_path": row["image_path"],
        "html_path": row.get("html_path", ""),
        "case_id": metadata["case_id"],
        "role": metadata["role"],
        "human_grade": int(metadata["human_grade"]),
        "human_agreement": metadata.get("human_agreement"),
        "human_reason": metadata.get("human_reason", ""),
        "sample_source": "v1_preserved",
    }


def build_manifest(
    annotations_path: Path,
    v1_manifest_path: Path,
) -> list[dict[str, Any]]:
    selected = [normalize_v1_row(row) for row in read_jsonl(v1_manifest_path)]
    selected_ids = {row["slide_id"] for row in selected}
    selected_cases = {row["case_id"] for row in selected}
    annotations = read_jsonl(annotations_path)

    for grade in (0, 1, 2):
        candidates = []
        for row in annotations:
            if row.get("human_median") != grade or row.get("slide_id") in selected_ids:
                continue
            image_path = SOURCE_ROOT / str(row.get("png_path") or "")
            if not row.get("case_id") or not image_path.is_file():
                continue
            candidates.append(row)
        candidates.sort(
            key=lambda row: (
                -float(row.get("agreement_rate") or 0),
                str(row["slide_id"]),
            )
        )
        added = 0
        for row in candidates:
            if row["case_id"] in selected_cases:
                continue
            selected.append(
                {
                    "slide_id": row["slide_id"],
                    "image_path": row["png_path"],
                    "html_path": row.get("html_path", ""),
                    "case_id": row["case_id"],
                    "role": row.get("role", "slide"),
                    "human_grade": int(row["human_median"]),
                    "human_agreement": row.get("agreement_rate"),
                    "human_reason": row.get("human_reasons", ""),
                    "sample_source": "v2_stratified_extension",
                }
            )
            selected_ids.add(row["slide_id"])
            selected_cases.add(row["case_id"])
            added += 1
            if added == 6:
                break
        if added != 6:
            raise ValueError(f"could not add six unique cases for grade {grade}")

    distribution = Counter(row["human_grade"] for row in selected)
    if distribution != Counter({0: 9, 1: 9, 2: 9, 3: 3}):
        raise ValueError(f"unexpected grade distribution: {distribution}")
    if len(selected) != 30 or len({row["slide_id"] for row in selected}) != 30:
        raise ValueError("v2 manifest must contain 30 unique slides")
    return selected


def flatten_dimension(dimension: str) -> list[dict[str, str]]:
    return [
        {
            "criterion_id": parent["id"],
            "criterion_label": parent["label"],
            **subquestion,
        }
        for parent in RUBRIC[dimension]
        for subquestion in parent["subquestions"]
    ]


def dimension_prompt(dimension: str) -> tuple[str, str, list[dict[str, str]]]:
    items = flatten_dimension(dimension)
    formatted = "\n".join(
        f"{index:02d}. ID={item['id']} | 大项={item['criterion_label']} | 子问题={item['label']} | {item['question']}"
        for index, item in enumerate(items, start=1)
    )
    dimension_name = "视觉美观" if dimension == "aesthetics" else "信息传达效果"
    instructions = f"""
你是严格的演示文稿评审员。只评估当前 slide 的{dimension_name}。

下面有 20 个互不替代的子问题。逐项给 1-5 整数分：
- 5：有明确可见的优秀证据，几乎无实质缺陷。
- 4：表现良好，仅有轻微问题。
- 3：基本可用，但存在明确摩擦。
- 2：有显著问题，影响专业展示或理解。
- 1：严重失败、不可读或无法恢复意图。

每项必须独立判断，其他项优点不能补偿本项。证据必须引用图片中可见事实，
用一句简洁中文表达；defects 最多两条。不要评价事实真伪。

{formatted}
""".strip()
    user_text = (
        "检查这张真实演示文稿 slide。必须恰好返回上述 20 个 subquestion_id，"
        "不得增加、删除、合并或重命名；先看缺陷，再给分。"
    )
    return instructions, user_text, items


def score_dimension(
    *,
    client: OpenAI,
    model: str,
    image_path: Path,
    slide_id: str,
    dimension: str,
    reasoning_effort: str,
    timeout_s: float,
    max_retries: int,
) -> dict[str, Any]:
    instructions, user_text, definitions = dimension_prompt(dimension)
    expected_ids = [item["id"] for item in definitions]
    last_error: Exception | None = None
    for attempt in range(max_retries + 1):
        try:
            response = client.responses.parse(
                model=model,
                instructions=instructions,
                input=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "input_text",
                                "text": f"Slide ID: {slide_id}\n\n{user_text}",
                            },
                            {
                                "type": "input_image",
                                "image_url": image_data_url(image_path),
                                "detail": "high",
                            },
                        ],
                    }
                ],
                reasoning={"effort": reasoning_effort},
                text_format=DimensionAssessment,
                max_output_tokens=9000,
                store=False,
                timeout=timeout_s,
            )
            parsed = response.output_parsed
            if parsed is None:
                raise RuntimeError("model response did not contain parsed output")
            by_id = {item.subquestion_id: item for item in parsed.results}
            if set(by_id) != set(expected_ids):
                raise ValueError("model subquestion IDs do not match rubric")
            usage = getattr(response, "usage", None)
            return {
                "results": [by_id[item_id].model_dump(mode="json") for item_id in expected_ids],
                "response_id": getattr(response, "id", None),
                "served_model": getattr(response, "model", model),
                "usage": usage.model_dump(mode="json") if hasattr(usage, "model_dump") else None,
                "prompt_sha256": sha256_bytes(
                    (instructions + "\n\n" + user_text).encode("utf-8")
                ),
            }
        except Exception as exc:
            last_error = exc
            if attempt >= max_retries:
                break
            time.sleep(2**attempt)
    raise RuntimeError(f"{slide_id}/{dimension} failed: {last_error}") from last_error


def criterion_rollups(results: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rollups = []
    for dimension in (
        "aesthetics",
        "content_accuracy",
        "communication_effectiveness",
    ):
        for parent in RUBRIC[dimension]:
            children = [row for row in results if row["criterion_id"] == parent["id"]]
            values = [row["score_1_5"] for row in children if row["score_1_5"] is not None]
            rollups.append(
                {
                    "criterion_id": parent["id"],
                    "criterion_label": parent["label"],
                    "dimension": dimension,
                    "score_1_5": round(sum(values) / len(values), 4) if values else None,
                    "source": (
                        "transparent_mean_of_direct_subquestions"
                        if values
                        else "abstention"
                    ),
                }
            )
    return rollups


def score_case(
    row: dict[str, Any],
    *,
    client: OpenAI,
    model: str,
    reasoning_effort: str,
    timeout_s: float,
    max_retries: int,
    asset_dir: Path,
) -> dict[str, Any]:
    source_image = SOURCE_ROOT / row["image_path"]
    target_image = asset_dir / f"{row['slide_id']}.png"
    asset_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_image, target_image)

    outputs = {
        dimension: score_dimension(
            client=client,
            model=model,
            image_path=source_image,
            slide_id=row["slide_id"],
            dimension=dimension,
            reasoning_effort=reasoning_effort,
            timeout_s=timeout_s,
            max_retries=max_retries,
        )
        for dimension in ("aesthetics", "communication_effectiveness")
    }

    results = []
    for dimension in (
        "aesthetics",
        "content_accuracy",
        "communication_effectiveness",
    ):
        definitions = flatten_dimension(dimension)
        if dimension == "content_accuracy":
            for definition in definitions:
                results.append(
                    {
                        "criterion_id": definition["criterion_id"],
                        "criterion_label": definition["criterion_label"],
                        "subquestion_id": definition["id"],
                        "subquestion_label": definition["label"],
                        "question": definition["question"],
                        "dimension": dimension,
                        "status": "not_assessable",
                        "source": "abstention",
                        "score_1_5": None,
                        "confidence_0_1": 0.0,
                        "evidence": ABSTENTION_REASON,
                        "defects": [],
                    }
                )
            continue
        scored = {
            item["subquestion_id"]: item for item in outputs[dimension]["results"]
        }
        for definition in definitions:
            assessment = scored[definition["id"]]
            results.append(
                {
                    "criterion_id": definition["criterion_id"],
                    "criterion_label": definition["criterion_label"],
                    "subquestion_id": definition["id"],
                    "subquestion_label": definition["label"],
                    "question": definition["question"],
                    "dimension": dimension,
                    "status": "scored",
                    "source": "model",
                    "score_1_5": assessment["score_1_5"],
                    "confidence_0_1": assessment["confidence_0_1"],
                    "evidence": assessment["evidence"],
                    "defects": assessment["defects"],
                }
            )

    return {
        "slide_id": row["slide_id"],
        "case_id": row["case_id"],
        "title": row["case_id"].replace("_", " "),
        "role": row["role"],
        "image": f"/benchmark-slides-v2/{target_image.name}",
        "image_sha256": sha256_bytes(target_image.read_bytes()),
        "human_aesthetics_grade_0_3": row["human_grade"],
        "human_agreement": row["human_agreement"],
        "human_reason": row["human_reason"],
        "sample_source": row["sample_source"],
        "results": results,
        "criterion_rollups": criterion_rollups(results),
        "calls": {
            dimension: {
                key: value
                for key, value in output.items()
                if key != "results"
            }
            for dimension, output in outputs.items()
        },
    }


def pairwise_accuracy(human: list[float], model: list[float]) -> float:
    correct = 0
    total = 0
    for left in range(len(human)):
        for right in range(left + 1, len(human)):
            if human[left] == human[right]:
                continue
            total += 1
            human_direction = human[left] > human[right]
            if model[left] == model[right]:
                correct += 0.5
            elif (model[left] > model[right]) == human_direction:
                correct += 1
    return correct / total


def summary(cases: list[dict[str, Any]]) -> dict[str, Any]:
    human = []
    model = []
    total_tokens = 0
    for case in cases:
        human.append(float(case["human_aesthetics_grade_0_3"]))
        scores = [
            row["score_1_5"]
            for row in case["results"]
            if row["dimension"] == "aesthetics"
        ]
        model.append(sum(scores) / len(scores))
        for call in case["calls"].values():
            total_tokens += int((call.get("usage") or {}).get("total_tokens") or 0)
    distribution = Counter(int(value) for value in human)
    means = {
        str(grade): round(
            sum(score for score, target in zip(model, human) if target == grade)
            / distribution[grade],
            4,
        )
        for grade in range(4)
    }
    return {
        "scored_subquestion_count": sum(
            row["status"] == "scored" for case in cases for row in case["results"]
        ),
        "abstention_count": sum(
            row["status"] == "not_assessable"
            for case in cases
            for row in case["results"]
        ),
        "total_tokens": total_tokens,
        "human_grade_distribution": {
            str(grade): distribution[grade] for grade in range(4)
        },
        "human_ai_aesthetics_spearman": round(spearman(human, model), 6),
        "human_ai_pairwise_accuracy": round(pairwise_accuracy(human, model), 6),
        "mean_aesthetics_by_human_grade": means,
        "scope": "Thirty-case stratified diagnostic; only three valid Human Grade 3 slides exist in the source annotations.",
    }


def write_jsonl(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        "".join(json.dumps(row, ensure_ascii=False) + "\n" for row in rows),
        encoding="utf-8",
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--annotations", type=Path, default=DEFAULT_ANNOTATIONS)
    parser.add_argument("--v1-manifest", type=Path, default=DEFAULT_V1_MANIFEST)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--asset-dir", type=Path, default=DEFAULT_ASSET_DIR)
    parser.add_argument("--run-dir", type=Path, default=DEFAULT_RUN_DIR)
    parser.add_argument("--base-url", default="http://127.0.0.1:19100/v1")
    parser.add_argument("--model", default="gpt-5.5")
    parser.add_argument("--reasoning-effort", default="medium")
    parser.add_argument("--workers", type=int, default=6)
    parser.add_argument("--timeout", type=float, default=300)
    parser.add_argument("--max-retries", type=int, default=2)
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()

    manifest = build_manifest(args.annotations, args.v1_manifest)
    manifest_path = args.run_dir / "manifest.jsonl"
    progress_path = args.run_dir / "progress.json"
    write_jsonl(manifest_path, manifest)

    rubric_hash = sha256_bytes(canonical_json(RUBRIC).encode("utf-8"))
    progress_metadata = {
        "model": args.model,
        "reasoning_effort": args.reasoning_effort,
        "rubric_sha256": rubric_hash,
        "manifest_sha256": sha256_bytes(manifest_path.read_bytes()),
    }
    progress = {"metadata": progress_metadata, "cases": {}}
    if progress_path.exists() and not args.force:
        progress = json.loads(progress_path.read_text(encoding="utf-8"))
        if progress.get("metadata") != progress_metadata:
            raise ValueError("progress metadata does not match this V2 run")
    completed = dict(progress.get("cases") or {})
    client = OpenAI(
        base_url=args.base_url.rstrip("/"),
        api_key="local-proxy",
        max_retries=0,
    )
    lock = threading.Lock()

    def run(row: dict[str, Any]) -> tuple[str, dict[str, Any]]:
        slide_id = row["slide_id"]
        if slide_id in completed and not args.force:
            return slide_id, completed[slide_id]
        result = score_case(
            row,
            client=client,
            model=args.model,
            reasoning_effort=args.reasoning_effort,
            timeout_s=args.timeout,
            max_retries=args.max_retries,
            asset_dir=args.asset_dir,
        )
        with lock:
            completed[slide_id] = result
            atomic_json(
                progress_path,
                {"metadata": progress_metadata, "cases": completed},
            )
        return slide_id, result

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        futures = {pool.submit(run, row): row["slide_id"] for row in manifest}
        for future in as_completed(futures):
            slide_id, _ = future.result()
            print(f"scored {slide_id} ({len(completed)}/30)", flush=True)

    atomic_json(
        progress_path,
        {"metadata": progress_metadata, "cases": completed},
    )
    ordered = [completed[row["slide_id"]] for row in manifest]
    artifact = {
        "experiment": {
            "id": "fine-grained-ppt-v2",
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "model": args.model,
            "provider": "local-responses-compatible",
            "reasoning_effort": args.reasoning_effort,
            "image_detail": "high",
            "rubric_version": RUBRIC_VERSION,
            "rubric_sha256": rubric_hash,
            "manifest_sha256": progress_metadata["manifest_sha256"],
            "scoring_mode": "two_dimension_subquestion_calls_per_slide",
            "case_count": 30,
            "model_call_count": 60,
            "subquestions_per_criterion": 4,
            "content_accuracy_policy": ABSTENTION_REASON,
        },
        "summary": summary(ordered),
        "rubric": RUBRIC,
        "cases": ordered,
    }
    atomic_json(args.output, artifact)
    print(f"wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
