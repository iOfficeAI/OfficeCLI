#!/usr/bin/env python3
"""Score 12 real benchmark slides on criterion-level PPT rubrics."""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import mimetypes
import os
import shutil
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openai import OpenAI
from pydantic import BaseModel, ConfigDict, Field, model_validator


REPO_ROOT = Path(__file__).resolve().parents[1]
SOURCE_ROOT = Path(
    os.environ.get(
        "OFFICE_REWARD_PPT_SOURCE_ROOT",
        REPO_ROOT / "source-documents" / "ppt",
    )
)
DEFAULT_MANIFEST = SOURCE_ROOT / "configs/scoring_case_manifest.jsonl"
DEFAULT_OUTPUT = REPO_ROOT / "app/rubric/fine-grained-results.json"
DEFAULT_ASSET_DIR = REPO_ROOT / "public/benchmark-slides"
DEFAULT_PROGRESS = REPO_ROOT / "experiments/fine-grained-ppt-v1/progress.json"
RUBRIC_VERSION = "ppt-criterion-rubric-v1"
ABSTENTION_REASON = "No authoritative reference evidence supplied."


CRITERIA: dict[str, list[dict[str, str]]] = {
    "aesthetics": [
        {
            "id": "ppt-layout",
            "label": "布局与构图",
            "question": "检查画面平衡、对齐、边距、分组，以及 16:9 画布是否被有意使用。",
            "score_5": "构图平衡、对齐精准、留白有目的，视觉重心明确且没有未完成区域。",
            "score_1": "构图失衡或破碎，存在大面积无意义空白、拥挤、错位或明显未完成区域。",
        },
        {
            "id": "ppt-typography",
            "label": "字体与层级",
            "question": "检查演示距离下的字号、行长、对比度、标题层级和字体一致性。",
            "score_5": "演示距离下完全易读，标题与正文层级稳定，字号和字重高度一致。",
            "score_1": "核心文字不可读、字号过小、层级混乱或字体使用严重不一致。",
        },
        {
            "id": "ppt-graphics",
            "label": "图表与图形",
            "question": "检查图表、图片、表格、图标、箭头和连接线的质量与整合程度。",
            "score_5": "图形清晰、标注完整、风格一致，并直接支持页面主要信息。",
            "score_1": "关键图形损坏、模糊、误导、难以辨认或与内容脱节。",
        },
        {
            "id": "ppt-technical",
            "label": "技术完整性",
            "question": "优先检查重叠、裁切、截断、溢出、断裂连接线和未完成区域。",
            "score_5": "没有可见重叠、裁切、溢出或断裂，所有元素均完整呈现。",
            "score_1": "存在影响理解的重叠、裁切、溢出、断裂或不可读核心内容。",
        },
        {
            "id": "ppt-economy",
            "label": "视觉经济性",
            "question": "检查是否同时避免无理由的大块空白和文档式拥挤。",
            "score_5": "信息密度恰当，留白帮助阅读，没有冗余装饰或拥挤区域。",
            "score_1": "页面极度空洞或拥挤，冗余内容明显妨碍视觉理解。",
        },
    ],
    "content_accuracy": [
        {"id": "ppt-central-claim", "label": "中心结论", "question": "核对页面中心结论。"},
        {"id": "ppt-numbers", "label": "数字与计算", "question": "核对数字、百分比、公式和图表数据。"},
        {"id": "ppt-labels", "label": "标签与来源", "question": "核对方法、数据集、基线、图例和来源。"},
        {"id": "ppt-scope", "label": "范围与因果", "question": "核对比较、因果、最高级和范围声明。"},
        {"id": "ppt-reference-coverage", "label": "证据覆盖", "question": "检查重要声明的参考证据覆盖。"},
    ],
    "communication_effectiveness": [
        {
            "id": "ppt-takeaway",
            "label": "核心信息清晰度",
            "question": "仅根据可见内容复述核心信息，并判断能否在数秒内稳定恢复。",
            "score_5": "数秒内即可准确复述唯一核心信息，不需要猜测或补充上下文。",
            "score_1": "无法恢复稳定核心信息，或多个冲突信息使页面不可理解。",
        },
        {
            "id": "ppt-alignment",
            "label": "标题与证据一致性",
            "question": "检查标题、强调、证据和结论是否共同支持同一信息。",
            "score_5": "标题、视觉强调、证据和结论完全同向，没有竞争性信息。",
            "score_1": "标题、强调和证据彼此矛盾或几乎没有关系。",
        },
        {
            "id": "ppt-reading-path",
            "label": "阅读路径",
            "question": "判断第一、第二和最终视觉落点是否形成连贯顺序。",
            "score_5": "阅读顺序自然且唯一，分组、对齐和连接关系无需解释。",
            "score_1": "阅读顺序混乱、循环或断裂，无法判断从哪里开始和结束。",
        },
        {
            "id": "ppt-comprehension",
            "label": "受众理解成本",
            "question": "检查术语、缩写、图例、密度和标签是否适合目标受众。",
            "score_5": "目标受众无需反复阅读即可理解，术语和图例解释充分。",
            "score_1": "核心内容依赖未解释术语或不可读细节，受众难以理解。",
        },
        {
            "id": "ppt-information-economy",
            "label": "信息取舍",
            "question": "检查每个内容块是否支持核心信息，是否重复或应移入备注。",
            "score_5": "每个元素都服务于结论，信息精炼且没有可见重复。",
            "score_1": "冗余、装饰或无关细节主导页面，核心信息被淹没。",
        },
    ],
}


class CriterionAssessment(BaseModel):
    model_config = ConfigDict(extra="forbid")

    criterion_id: str
    score_1_5: int = Field(ge=1, le=5)
    confidence_0_1: float = Field(ge=0, le=1)
    evidence: str = Field(min_length=20)
    defects: list[str] = Field(default_factory=list, max_length=4)


class DimensionAssessment(BaseModel):
    model_config = ConfigDict(extra="forbid")

    results: list[CriterionAssessment] = Field(min_length=5, max_length=5)

    @model_validator(mode="after")
    def unique_criteria(self) -> "DimensionAssessment":
        identifiers = [item.criterion_id for item in self.results]
        if len(set(identifiers)) != len(identifiers):
            raise ValueError("criterion IDs must be unique")
        return self


def canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def atomic_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def rankdata(values: list[float]) -> list[float]:
    order = sorted(range(len(values)), key=values.__getitem__)
    ranks = [0.0] * len(values)
    index = 0
    while index < len(order):
        end = index + 1
        while end < len(order) and values[order[end]] == values[order[index]]:
            end += 1
        average_rank = (index + end - 1) / 2.0 + 1.0
        for cursor in range(index, end):
            ranks[order[cursor]] = average_rank
        index = end
    return ranks


def spearman(xs: list[float], ys: list[float]) -> float:
    ranked_x = rankdata(xs)
    ranked_y = rankdata(ys)
    mean_x = sum(ranked_x) / len(ranked_x)
    mean_y = sum(ranked_y) / len(ranked_y)
    numerator = sum((x - mean_x) * (y - mean_y) for x, y in zip(ranked_x, ranked_y))
    denominator_x = sum((x - mean_x) ** 2 for x in ranked_x) ** 0.5
    denominator_y = sum((y - mean_y) ** 2 for y in ranked_y) ** 0.5
    return numerator / (denominator_x * denominator_y)


def experiment_summary(cases: list[dict[str, Any]]) -> dict[str, Any]:
    aesthetics_means = []
    human_grades = []
    scored_count = 0
    abstention_count = 0
    total_tokens = 0
    for case in cases:
        aesthetics = [
            result["score_1_5"]
            for result in case["results"]
            if result["dimension"] == "aesthetics" and result["status"] == "scored"
        ]
        aesthetics_means.append(sum(aesthetics) / len(aesthetics))
        human_grades.append(float(case["human_aesthetics_grade_0_3"]))
        scored_count += sum(result["status"] == "scored" for result in case["results"])
        abstention_count += sum(
            result["status"] == "not_assessable" for result in case["results"]
        )
        for call in case["calls"].values():
            usage = call.get("usage") or {}
            total_tokens += int(usage.get("total_tokens") or 0)

    means_by_grade = {}
    for grade in range(4):
        values = [
            score
            for score, human_grade in zip(aesthetics_means, human_grades)
            if human_grade == grade
        ]
        means_by_grade[str(grade)] = round(sum(values) / len(values), 4)
    return {
        "scored_criterion_count": scored_count,
        "abstention_count": abstention_count,
        "total_tokens": total_tokens,
        "human_ai_aesthetics_spearman": round(
            spearman(human_grades, aesthetics_means),
            6,
        ),
        "mean_aesthetics_by_human_grade": means_by_grade,
        "scope": "Preliminary 12-case diagnostic; criterion-level human labels are not yet available.",
    }


def image_data_url(path: Path) -> str:
    media_type = mimetypes.guess_type(path.name)[0] or "image/png"
    return f"data:{media_type};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"


def dimension_prompt(dimension: str) -> tuple[str, str]:
    criteria = CRITERIA[dimension]
    criteria_text = "\n\n".join(
        (
            f"ID: {item['id']}\n名称: {item['label']}\n检查: {item['question']}\n"
            f"5分锚点: {item['score_5']}\n1分锚点: {item['score_1']}"
        )
        for item in criteria
    )
    instructions = f"""
你是严格的演示文稿评审员。只评估当前 slide 的{('视觉美观' if dimension == 'aesthetics' else '信息传达效果')}。

对下面五个小项分别给 1-5 的整数分。每项必须独立判断，其他项的优点不能补偿本项缺陷。
先识别可见缺陷，再选择分数；不确定时取较低分。5 分必须有明确正面证据，1 分表示严重失败。
只依据图片中可见内容。证据和缺陷使用简洁中文，不要评价事实真伪。

评分小项：
{criteria_text}
""".strip()
    user_text = (
        "检查这张真实演示文稿 slide。必须恰好返回上述五个 criterion_id，"
        "不得增加、删除或重命名。每个 evidence 必须引用该图中可见的具体证据。"
    )
    return instructions, user_text


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
    instructions, user_text = dimension_prompt(dimension)
    expected_ids = [item["id"] for item in CRITERIA[dimension]]
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
                            {"type": "input_text", "text": f"Slide ID: {slide_id}\n\n{user_text}"},
                            {"type": "input_image", "image_url": image_data_url(image_path), "detail": "high"},
                        ],
                    }
                ],
                reasoning={"effort": reasoning_effort},
                text_format=DimensionAssessment,
                max_output_tokens=3200,
                store=False,
                timeout=timeout_s,
            )
            parsed = response.output_parsed
            if parsed is None:
                raise RuntimeError("model response did not contain parsed output")
            by_id = {item.criterion_id: item for item in parsed.results}
            if set(by_id) != set(expected_ids):
                raise ValueError(f"criterion mismatch: {sorted(by_id)} != {sorted(expected_ids)}")
            ordered = [by_id[criterion_id].model_dump(mode="json") for criterion_id in expected_ids]
            usage = getattr(response, "usage", None)
            return {
                "results": ordered,
                "response_id": getattr(response, "id", None),
                "served_model": getattr(response, "model", model),
                "usage": usage.model_dump(mode="json") if hasattr(usage, "model_dump") else None,
                "prompt_sha256": sha256_bytes((instructions + "\n\n" + user_text).encode("utf-8")),
            }
        except Exception as exc:  # retry malformed or transient proxy responses
            last_error = exc
            if attempt >= max_retries:
                break
            time.sleep(2**attempt)
    raise RuntimeError(f"{slide_id}/{dimension} failed: {last_error}") from last_error


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
    if not source_image.is_file():
        raise FileNotFoundError(source_image)
    asset_dir.mkdir(parents=True, exist_ok=True)
    target_image = asset_dir / f"{row['slide_id']}.png"
    shutil.copy2(source_image, target_image)
    image_hash = sha256_bytes(target_image.read_bytes())

    dimension_outputs = {}
    for dimension in ("aesthetics", "communication_effectiveness"):
        dimension_outputs[dimension] = score_dimension(
            client=client,
            model=model,
            image_path=source_image,
            slide_id=row["slide_id"],
            dimension=dimension,
            reasoning_effort=reasoning_effort,
            timeout_s=timeout_s,
            max_retries=max_retries,
        )

    results = []
    for dimension in ("aesthetics", "content_accuracy", "communication_effectiveness"):
        if dimension == "content_accuracy":
            for criterion in CRITERIA[dimension]:
                results.append(
                    {
                        "criterion_id": criterion["id"],
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
        for assessment in dimension_outputs[dimension]["results"]:
            results.append(
                {
                    "criterion_id": assessment["criterion_id"],
                    "dimension": dimension,
                    "status": "scored",
                    "source": "model",
                    "score_1_5": assessment["score_1_5"],
                    "confidence_0_1": assessment["confidence_0_1"],
                    "evidence": assessment["evidence"],
                    "defects": assessment["defects"],
                }
            )

    metadata = row.get("metadata") or {}
    return {
        "slide_id": row["slide_id"],
        "case_id": metadata.get("case_id", row["slide_id"]),
        "title": str(metadata.get("case_id", row["slide_id"])).replace("_", " "),
        "role": metadata.get("role", "slide"),
        "image": f"/benchmark-slides/{target_image.name}",
        "image_sha256": image_hash,
        "human_aesthetics_grade_0_3": metadata["human_grade"],
        "human_agreement": metadata.get("human_agreement"),
        "human_reason": metadata.get("human_reason", ""),
        "results": results,
        "calls": {
            dimension: {
                key: value
                for key, value in dimension_outputs[dimension].items()
                if key != "results"
            }
            for dimension in dimension_outputs
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, default=DEFAULT_MANIFEST)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--asset-dir", type=Path, default=DEFAULT_ASSET_DIR)
    parser.add_argument("--progress", type=Path, default=DEFAULT_PROGRESS)
    parser.add_argument("--base-url", default="http://127.0.0.1:19100/v1")
    parser.add_argument("--model", default="gpt-5.6-sol")
    parser.add_argument("--reasoning-effort", default="medium")
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--timeout", type=float, default=180)
    parser.add_argument("--max-retries", type=int, default=2)
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()

    rows = read_jsonl(args.manifest)
    if len(rows) != 12:
        raise ValueError(f"expected 12 benchmark rows, found {len(rows)}")
    grades = [int((row.get("metadata") or {})["human_grade"]) for row in rows]
    if {grade: grades.count(grade) for grade in set(grades)} != {0: 3, 1: 3, 2: 3, 3: 3}:
        raise ValueError("benchmark must contain three cases per human grade")

    rubric_hash = sha256_bytes(canonical_json(CRITERIA).encode("utf-8"))
    progress_metadata = {
        "model": args.model,
        "reasoning_effort": args.reasoning_effort,
        "rubric_sha256": rubric_hash,
    }
    progress: dict[str, Any] = {"metadata": progress_metadata, "cases": {}}
    if args.progress.exists() and not args.force:
        progress = json.loads(args.progress.read_text(encoding="utf-8"))
        existing_metadata = progress.get("metadata")
        if existing_metadata is not None and existing_metadata != progress_metadata:
            raise ValueError("progress metadata does not match this scoring run")
    completed = dict(progress.get("cases") or {})

    client = OpenAI(base_url=args.base_url.rstrip("/"), api_key="local-proxy", max_retries=0)
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
                args.progress,
                {"metadata": progress_metadata, "cases": completed},
            )
        return slide_id, result

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        futures = {pool.submit(run, row): row["slide_id"] for row in rows}
        for future in as_completed(futures):
            slide_id, _result = future.result()
            print(f"scored {slide_id} ({len(completed)}/12)", flush=True)

    atomic_json(
        args.progress,
        {"metadata": progress_metadata, "cases": completed},
    )

    ordered_cases = [completed[row["slide_id"]] for row in rows]
    artifact = {
        "experiment": {
            "id": "fine-grained-ppt-v1",
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "model": args.model,
            "provider": "local-responses-compatible",
            "reasoning_effort": args.reasoning_effort,
            "image_detail": "high",
            "rubric_version": RUBRIC_VERSION,
            "rubric_sha256": rubric_hash,
            "scoring_mode": "two_dimension_calls_per_slide",
            "case_count": len(ordered_cases),
            "model_call_count": len(ordered_cases) * 2,
            "content_accuracy_policy": ABSTENTION_REASON,
        },
        "summary": experiment_summary(ordered_cases),
        "criteria": CRITERIA,
        "cases": ordered_cases,
    }
    atomic_json(args.output, artifact)
    print(f"wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
