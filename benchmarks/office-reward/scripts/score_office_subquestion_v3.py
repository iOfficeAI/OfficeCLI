#!/usr/bin/env python3
"""Render and score Word/Excel units, then combine them with PPT V2."""

from __future__ import annotations

import argparse
import copy
import json
import os
import subprocess
import threading
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openai import OpenAI
from PIL import Image, ImageStat
from pydantic import BaseModel, ConfigDict, Field, model_validator

from score_fine_grained_ppt import (
    ABSTENTION_REASON,
    REPO_ROOT,
    atomic_json,
    canonical_json,
    image_data_url,
    sha256_bytes,
)
from score_subquestion_ppt_v2 import RUBRIC as PPT_RUBRIC, criterion


DEFAULT_V2 = REPO_ROOT / "app/rubric/subquestion-results-v2.json"
DEFAULT_OUTPUT = REPO_ROOT / "app/rubric/office-subquestion-results-v3.json"
DEFAULT_ASSET_DIR = REPO_ROOT / "public/benchmark-units-v3"
DEFAULT_RUN_DIR = REPO_ROOT / "experiments/office-subquestion-v3"
RUBRIC_VERSION = "office-subquestion-rubric-v3"
OFFICECLI_ROOT = REPO_ROOT.parents[1]
DEFAULT_CORPUS_ROOT = Path(
    os.environ.get(
        "OFFICE_REWARD_CORPUS_ROOT",
        REPO_ROOT / "source-documents",
    )
)


DOCUMENT_CASES: list[dict[str, str]] = [
    {
        "case_uid": "docx-officecli-document-formatting",
        "format": "docx",
        "source": "repo://examples/word/document-formatting.docx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "docx-officecli-charts",
        "format": "docx",
        "source": "repo://examples/word/charts.docx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "docx-officecli-diagram",
        "format": "docx",
        "source": "repo://examples/word/diagram.docx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "docx-officecli-formulas",
        "format": "docx",
        "source": "repo://examples/word/formulas.docx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "docx-officecli-numbering",
        "format": "docx",
        "source": "repo://examples/word/numbering.docx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "docx-officecli-pictures",
        "format": "docx",
        "source": "repo://examples/word/pictures.docx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "docx-officecli-sections",
        "format": "docx",
        "source": "repo://examples/word/sections.docx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "docx-officecli-tables",
        "format": "docx",
        "source": "repo://examples/word/tables.docx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "docx-real-dog-report",
        "format": "docx",
        "source": "corpus://docx/机器狗巡检系统测试关卡集合_大作业报告_完善版.docx",
        "source_set": "real report",
    },
    {
        "case_uid": "docx-real-prithvi-report",
        "format": "docx",
        "source": "corpus://docx/prithvi_research_report.docx",
        "source_set": "real report",
    },
    {
        "case_uid": "docx-real-chinese-report",
        "format": "docx",
        "source": "corpus://docx/chinese_report_full.docx",
        "source_set": "real report",
    },
    {
        "case_uid": "docx-real-graph-theory",
        "format": "docx",
        "source": "corpus://docx/图论问题.docx",
        "source_set": "real report",
    },
    {
        "case_uid": "xlsx-real-construction-progress",
        "format": "xlsx",
        "source": "corpus://xlsx/construction_site_progress.xlsx",
        "source_set": "real operational workbook",
    },
    {
        "case_uid": "xlsx-real-fintech-fraud",
        "format": "xlsx",
        "source": "corpus://xlsx/fintech_transaction_fraud.xlsx",
        "source_set": "real operational workbook",
    },
    {
        "case_uid": "xlsx-real-financial-dashboard",
        "format": "xlsx",
        "source": "corpus://xlsx/executive_financial_dashboard.xlsx",
        "source_set": "real operational workbook",
    },
    {
        "case_uid": "xlsx-officecli-basic-charts",
        "format": "xlsx",
        "source": "repo://examples/excel/charts/charts-basic.xlsx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "xlsx-real-hotel-revenue",
        "format": "xlsx",
        "source": "corpus://xlsx/hotel_bookings_revenue_ops.xlsx",
        "source_set": "real operational workbook",
    },
    {
        "case_uid": "xlsx-officecli-sparklines",
        "format": "xlsx",
        "source": "repo://examples/excel/sparklines.xlsx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "xlsx-real-construction-gantt",
        "format": "xlsx",
        "source": "corpus://xlsx/construction_gantt_control.xlsx",
        "source_set": "real operational workbook",
    },
    {
        "case_uid": "xlsx-real-saas-failure",
        "format": "xlsx",
        "source": "corpus://xlsx/saas_revenue_noskills_fail.xlsx",
        "source_set": "real operational workbook",
    },
    {
        "case_uid": "xlsx-officecli-cell-formatting",
        "format": "xlsx",
        "source": "repo://examples/excel/cell-formatting.xlsx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "xlsx-officecli-advanced-charts",
        "format": "xlsx",
        "source": "repo://examples/excel/charts/charts-advanced.xlsx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "xlsx-officecli-conditional-formatting",
        "format": "xlsx",
        "source": "repo://examples/excel/conditional-formatting.xlsx",
        "source_set": "OfficeCLI feature example",
    },
    {
        "case_uid": "xlsx-officecli-pivot-tables",
        "format": "xlsx",
        "source": "repo://examples/excel/pivot-tables.xlsx",
        "source_set": "OfficeCLI feature example",
    },
]


WORD_RUBRIC: dict[str, list[dict[str, Any]]] = {
    "aesthetics": [
        criterion("word-page-composition", "页面构图", [
            ("margins", "页边距", "正文区域和页边距是否稳定、均衡且适合阅读？"),
            ("balance", "页面平衡", "页面上下与左右的内容重量是否平衡？"),
            ("alignment", "对象对齐", "段落、标题、表格和图片是否沿一致参考线对齐？"),
            ("pagination", "分页观感", "当前页面是否呈现自然边界，避免突兀空白或拥堵页脚？"),
        ]),
        criterion("word-typography", "字体与标题层级", [
            ("body", "正文可读性", "正文的字号、行长、行距和对比度是否舒适？"),
            ("headings", "标题层级", "标题级别是否通过字号、字重和间距清楚区分？"),
            ("styles", "样式一致性", "同类段落和标题是否使用一致的字体与格式？"),
            ("emphasis", "强调克制", "粗体、斜体、颜色和高亮是否准确且不过度？"),
        ]),
        criterion("word-spacing", "间距与阅读节奏", [
            ("paragraphs", "段落间距", "段落之间是否有稳定、可预测的垂直节奏？"),
            ("headings", "标题间距", "标题与前后正文的距离是否清楚表达结构？"),
            ("lists", "列表节奏", "列表缩进、项目间距和编号是否整齐易扫读？"),
            ("objects", "对象间距", "表格、图片、图表与正文之间是否有足够空间？"),
        ]),
        criterion("word-media", "表格图表与图片", [
            ("legibility", "对象可读性", "表格、图表、图片和公式是否清晰可辨？"),
            ("placement", "对象位置", "对象是否靠近相关正文且没有破坏阅读顺序？"),
            ("captions", "题注与标签", "题注、编号、图例和表头是否清楚完整？"),
            ("integration", "环绕与整合", "对象尺寸、对齐和文字环绕是否与页面协调？"),
        ]),
        criterion("word-technical", "技术完整性", [
            ("clipping", "裁切溢出", "文字或对象是否被裁切、截断或越出页面？"),
            ("orphans", "孤立结构", "是否存在孤立标题、落单题注或异常分页片段？"),
            ("headers", "页眉页脚", "页眉、页脚和页码是否完整、克制且不干扰正文？"),
            ("artifacts", "渲染故障", "是否存在乱码、缺图、重叠或明显未完成格式？"),
        ]),
    ],
    "content_accuracy": [
        criterion("word-claims", "正文与标题声明", [
            ("central", "中心声明", "中心事实性声明是否与参考证据一致？"),
            ("headings", "标题声明", "标题和摘要是否准确概括正文？"),
            ("tables", "表格图表声明", "表格、图表和题注中的声明是否准确？"),
            ("conclusions", "结论对应", "结论和建议是否得到正文证据支持？"),
        ]),
        criterion("word-values", "数值单位与日期", [
            ("numbers", "数值", "重要数值是否与参考材料一致？"),
            ("units", "单位", "单位、精度和符号是否准确？"),
            ("dates", "日期", "日期和时间范围是否正确？"),
            ("names", "名称", "人名、机构和专有名词是否准确？"),
        ]),
        criterion("word-cross-references", "交叉引用", [
            ("figures", "图表引用", "正文是否引用正确的图表和题注？"),
            ("sections", "章节引用", "章节编号和内部引用是否正确？"),
            ("notes", "脚注尾注", "脚注、尾注和标记是否对应正确内容？"),
            ("links", "链接目标", "可见链接和引用目标是否准确？"),
        ]),
        criterion("word-citations", "引用与归因", [
            ("identity", "来源身份", "引用是否指向正确作者或来源？"),
            ("support", "支撑关系", "引用是否实际支持邻近声明？"),
            ("placement", "引用位置", "引用是否放在其所支持内容附近？"),
            ("coverage", "引用覆盖", "重要外部声明是否都有来源？"),
        ]),
        criterion("word-reference-coverage", "证据覆盖", [
            ("central", "核心覆盖", "参考材料是否覆盖中心论点？"),
            ("numbers", "数字覆盖", "参考材料是否覆盖关键数值？"),
            ("supporting", "支撑覆盖", "重要支撑声明是否可核验？"),
            ("residual", "剩余未知", "不可核验内容是否被明确区分？"),
        ]),
    ],
    "communication_effectiveness": [
        criterion("word-purpose", "目的清晰度", [
            ("objective", "文档目标", "当前页面能否让读者识别文档目标？"),
            ("audience", "目标读者", "语气和信息深度是否暗示清楚的目标读者？"),
            ("outcome", "预期结果", "读者是否知道应理解、决定或执行什么？"),
            ("opening", "开篇效率", "页面开头是否快速建立主题与上下文？"),
        ]),
        criterion("word-progression", "章节推进", [
            ("sequence", "逻辑顺序", "标题和段落是否按自然顺序推进？"),
            ("transitions", "段落过渡", "相邻段落之间的关系是否清楚？"),
            ("evidence", "论据位置", "证据是否紧邻其支持的主张？"),
            ("closure", "页面收束", "页面是否自然收束到结论或下一步？"),
        ]),
        criterion("word-navigation", "导航能力", [
            ("headings", "标题导航", "标题层级是否帮助快速定位？"),
            ("lists", "列表导航", "列表和编号是否帮助扫描步骤或要点？"),
            ("captions", "对象导航", "题注和交叉引用是否帮助关联正文与对象？"),
            ("running", "页眉页脚导航", "页码和页眉页脚是否提供有效位置感？"),
        ]),
        criterion("word-density", "冗余与密度", [
            ("density", "页面密度", "页面信息量是否适中？"),
            ("sentences", "句子负担", "句子长度和复杂度是否适合连续阅读？"),
            ("repetition", "内容重复", "是否避免重复陈述同一观点？"),
            ("whitespace", "阅读留白", "留白是否支持段落和对象分组？"),
        ]),
        criterion("word-audience-fit", "受众适配", [
            ("terms", "术语解释", "专业术语和缩写是否得到解释？"),
            ("context", "背景充分性", "读者理解本页所需背景是否齐全？"),
            ("detail", "细节深度", "细节层级是否适合目标读者？"),
            ("summary", "摘要质量", "页面是否突出最需要记住的信息？"),
        ]),
    ],
}


EXCEL_RUBRIC: dict[str, list[dict[str, Any]]] = {
    "aesthetics": [
        criterion("excel-used-range", "有效区域布局", [
            ("compactness", "区域紧凑性", "已填充区域是否紧凑而不散落？"),
            ("alignment", "网格对齐", "数据区、摘要区和图表是否沿稳定网格对齐？"),
            ("spacing", "区块间距", "不同功能区之间是否有清楚且适量的间距？"),
            ("canvas", "画布利用", "是否避免大面积无意义空白或内容挤在角落？"),
        ]),
        criterion("excel-sizing-formats", "尺寸与数字格式", [
            ("columns", "列宽", "列宽是否避免截断和过度换行？"),
            ("rows", "行高", "行高是否适配内容并保持一致？"),
            ("numbers", "数字格式", "货币、百分比、日期和小数格式是否一致？"),
            ("readability", "单元格可读性", "文字和数值在正常缩放下是否清楚？"),
        ]),
        criterion("excel-hierarchy", "输入计算输出层级", [
            ("inputs", "输入识别", "可编辑输入是否通过位置或样式清楚识别？"),
            ("calculations", "计算区识别", "中间计算是否与输入和结果区分？"),
            ("outputs", "结果强调", "关键输出和 KPI 是否得到适当强调？"),
            ("consistency", "样式语义", "同一颜色和样式是否始终表达同一含义？"),
        ]),
        criterion("excel-charts-formatting", "图表与条件格式", [
            ("placement", "图表位置", "图表是否与相关数据相邻且不遮挡内容？"),
            ("labels", "图表标签", "标题、坐标轴、图例和数据标签是否完整？"),
            ("legibility", "图表可读性", "图表尺寸、颜色和比例是否便于比较？"),
            ("conditional", "条件格式", "条件格式是否克制、可解释且没有误导？"),
        ]),
        criterion("excel-scanability", "扫描效率与完整性", [
            ("focal-order", "焦点顺序", "视线是否自然从摘要到趋势再到明细？"),
            ("exceptions", "异常突出", "异常和需要行动的状态是否清楚突出？"),
            ("clipping", "裁切重叠", "是否存在截断、重叠或对象越界？"),
            ("finish", "完成度", "工作表是否避免原始数据倾倒或未完成区域？"),
        ]),
    ],
    "content_accuracy": [
        criterion("excel-values", "值单位与汇总", [
            ("values", "关键值", "关键可见值是否与参考数据一致？"),
            ("units", "单位格式", "单位、符号和精度是否准确？"),
            ("totals", "合计小计", "合计、小计和比例是否正确？"),
            ("dates", "日期期间", "日期、期间和时间轴是否准确？"),
        ]),
        criterion("excel-formulas", "公式与计算", [
            ("logic", "公式逻辑", "核心公式逻辑是否正确？"),
            ("references", "公式引用", "单元格和跨表引用是否正确？"),
            ("cached", "缓存结果", "显示值是否与公式计算结果一致？"),
            ("errors", "错误处理", "公式错误和缺失值是否正确处理？"),
        ]),
        criterion("excel-chart-sources", "图表数据源", [
            ("ranges", "源范围", "图表源范围是否指向正确数据？"),
            ("series", "系列映射", "系列、类别和标签是否正确映射？"),
            ("values", "图表数值", "图表展示是否与源单元格一致？"),
            ("coverage", "数据覆盖", "图表是否遗漏或重复重要数据？"),
        ]),
        criterion("excel-references", "命名与跨表引用", [
            ("named", "命名范围", "命名范围是否有效且指向正确区域？"),
            ("cross-sheet", "跨表引用", "跨工作表引用是否准确？"),
            ("hidden", "隐藏依赖", "隐藏行列或工作表依赖是否合理？"),
            ("external", "外部链接", "外部引用是否有效并有清楚来源？"),
        ]),
        criterion("excel-assumptions", "假设与证据覆盖", [
            ("assumptions", "业务假设", "重要业务假设是否明确且有依据？"),
            ("units", "单位一致性", "输入、计算和输出单位是否一致？"),
            ("claims", "结论覆盖", "可见结论是否得到数据支持？"),
            ("residual", "未知项", "不可核验值和假设是否被清楚区分？"),
        ]),
    ],
    "communication_effectiveness": [
        criterion("excel-workflow", "输入到输出工作流", [
            ("inputs", "输入入口", "用户能否快速识别需要填写或调整的输入？"),
            ("calculations", "计算路径", "输入如何转化为结果是否容易追踪？"),
            ("outputs", "输出识别", "关键输出和 KPI 是否一眼可见？"),
            ("exceptions", "异常行动", "异常是否清楚连接到需要采取的行动？"),
        ]),
        criterion("excel-labeling", "标签单位与图例", [
            ("headers", "表头标签", "行列和区块标签是否明确？"),
            ("units", "单位说明", "数值单位和期间是否无需猜测？"),
            ("legends", "图例说明", "颜色、图标和状态图例是否自解释？"),
            ("notes", "说明文字", "必要假设和口径是否有简洁说明？"),
        ]),
        criterion("excel-context", "层级与工作表上下文", [
            ("summary", "摘要到明细", "摘要与明细区域的关系是否清楚？"),
            ("navigation", "表内导航", "冻结窗格、分组和位置是否帮助定位？"),
            ("sheet", "工作表身份", "当前工作表的用途和范围是否明确？"),
            ("hierarchy", "视觉层级", "标题、区块、表头和数据层级是否稳定？"),
        ]),
        criterion("excel-explainability", "假设与公式可解释性", [
            ("assumptions", "假设可见性", "关键假设是否清楚标识？"),
            ("formulas", "公式说明", "关键计算是否有足够标签或注释？"),
            ("exceptions", "例外说明", "错误、缺失和例外情况是否解释？"),
            ("audit", "可追踪性", "读者能否从结果追溯到相关输入和计算？"),
        ]),
        criterion("excel-decision", "决策相关性", [
            ("takeaway", "核心结论", "工作表是否传达稳定的核心结论？"),
            ("priorities", "优先级", "重要指标和异常是否按决策优先级呈现？"),
            ("actions", "行动指向", "输出是否清楚指向下一步行动？"),
            ("economy", "信息取舍", "是否避免与决策无关的重复或噪声？"),
        ]),
    ],
}


RUBRIC_BY_FORMAT = {
    "pptx": PPT_RUBRIC,
    "docx": WORD_RUBRIC,
    "xlsx": EXCEL_RUBRIC,
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
        if len({item.subquestion_id for item in self.results}) != 20:
            raise ValueError("subquestion IDs must contain 20 unique values")
        return self


def flatten(rubric: dict[str, list[dict[str, Any]]], dimension: str) -> list[dict[str, str]]:
    return [
        {
            "criterion_id": parent["id"],
            "criterion_label": parent["label"],
            **subquestion,
        }
        for parent in rubric[dimension]
        for subquestion in parent["subquestions"]
    ]


def resolve_source(source_uri: str, corpus_root: Path) -> Path:
    if source_uri.startswith("repo://"):
        return OFFICECLI_ROOT / source_uri.removeprefix("repo://")
    if source_uri.startswith("corpus://"):
        return corpus_root / source_uri.removeprefix("corpus://")
    raise ValueError(f"unsupported source URI: {source_uri}")


def document_manifest(corpus_root: Path) -> list[dict[str, Any]]:
    rows = []
    for item in DOCUMENT_CASES:
        source = resolve_source(item["source"], corpus_root)
        if not source.is_file():
            raise FileNotFoundError(source)
        rows.append(
            {
                **item,
                "_source_path": source,
                "source_name": source.name,
                "source_document_sha256": sha256_bytes(source.read_bytes()),
                "unit_type": "page" if item["format"] == "docx" else "sheet",
                "unit_name": "page 1" if item["format"] == "docx" else "first visible sheet",
            }
        )
    counts = Counter(row["format"] for row in rows)
    if counts != Counter({"docx": 12, "xlsx": 12}):
        raise ValueError(f"unexpected document counts: {counts}")
    return rows


def render_unit(row: dict[str, Any], output: Path) -> None:
    if output.exists():
        try:
            validate_image(output)
            return
        except ValueError:
            pass
    output.parent.mkdir(parents=True, exist_ok=True)
    args = [
        "officecli",
        "view",
        str(row["_source_path"]),
        "screenshot",
        "--screenshot-width",
        "1200" if row["format"] == "docx" else "1600",
        "--screenshot-height",
        "1600" if row["format"] == "docx" else "1200",
        "-o",
        str(output),
    ]
    if row["format"] == "docx":
        args.extend(["--page", "1", "--render", "html"])
    env = dict(os.environ)
    env["OFFICECLI_SKIP_UPDATE"] = "1"
    env["OFFICECLI_NO_AUTO_RESIDENT"] = "1"
    result = subprocess.run(
        args,
        capture_output=True,
        text=True,
        timeout=180,
        env=env,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"OfficeCLI render failed for {row['case_uid']}: "
            f"{(result.stderr or result.stdout)[-500:]}"
        )
    validate_image(output)


def validate_image(path: Path) -> None:
    if not path.is_file() or path.stat().st_size <= 10_000:
        raise ValueError(f"missing or undersized render: {path}")
    with Image.open(path) as image:
        image.verify()
    with Image.open(path).convert("RGB").resize((160, 120)) as image:
        if max(ImageStat.Stat(image).stddev) < 1.0:
            raise ValueError(f"blank render: {path}")


def dimension_prompt(format_name: str, dimension: str) -> tuple[str, str, list[dict[str, str]]]:
    rubric = RUBRIC_BY_FORMAT[format_name]
    items = flatten(rubric, dimension)
    medium = "Word 页面" if format_name == "docx" else "Excel 工作表"
    dimension_name = "视觉美观" if dimension == "aesthetics" else "信息传达效果"
    formatted = "\n".join(
        f"{index:02d}. ID={item['id']} | 板块={item['criterion_label']} | 子问题={item['label']} | {item['question']}"
        for index, item in enumerate(items, start=1)
    )
    instructions = f"""
你是严格的 Office 文档评审员。只评估当前{medium}的{dimension_name}。
下面 20 个子问题互不替代。逐项给 1-5 整数分：5=优秀且有明确可见证据；
4=良好仅有轻微问题；3=可用但有明确摩擦；2=显著问题；1=严重失败。
每项独立判断，其他项优点不能补偿。证据必须引用图片中的具体可见事实，
用一句简洁中文表达；defects 最多两条。不要评价事实真伪。

{formatted}
""".strip()
    user_text = (
        f"检查这张真实{medium}截图。必须恰好返回上述 20 个 subquestion_id，"
        "不得增加、删除、合并或重命名。"
    )
    return instructions, user_text, items


def score_dimension(
    *,
    client: OpenAI,
    model: str,
    image_path: Path,
    case_uid: str,
    format_name: str,
    dimension: str,
    reasoning_effort: str,
    timeout_s: float,
    max_retries: int,
) -> dict[str, Any]:
    instructions, user_text, definitions = dimension_prompt(format_name, dimension)
    expected_ids = [item["id"] for item in definitions]
    last_error: Exception | None = None
    for attempt in range(max_retries + 1):
        try:
            response = client.responses.parse(
                model=model,
                instructions=instructions,
                input=[{
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": f"Case ID: {case_uid}\n\n{user_text}"},
                        {"type": "input_image", "image_url": image_data_url(image_path), "detail": "high"},
                    ],
                }],
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
                "prompt_sha256": sha256_bytes((instructions + "\n\n" + user_text).encode("utf-8")),
            }
        except Exception as exc:
            last_error = exc
            if attempt >= max_retries:
                break
            time.sleep(2**attempt)
    raise RuntimeError(f"{case_uid}/{dimension} failed: {last_error}") from last_error


def criterion_rollups(
    rubric: dict[str, list[dict[str, Any]]],
    results: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    rollups = []
    for dimension in ("aesthetics", "content_accuracy", "communication_effectiveness"):
        for parent in rubric[dimension]:
            values = [
                row["score_1_5"]
                for row in results
                if row["criterion_id"] == parent["id"] and row["score_1_5"] is not None
            ]
            rollups.append({
                "criterion_id": parent["id"],
                "criterion_label": parent["label"],
                "dimension": dimension,
                "score_1_5": round(sum(values) / len(values), 4) if values else None,
                "source": "transparent_mean_of_direct_subquestions" if values else "abstention",
            })
    return rollups


def score_document_case(
    row: dict[str, Any],
    *,
    client: OpenAI,
    model: str,
    reasoning_effort: str,
    timeout_s: float,
    max_retries: int,
    asset_dir: Path,
    officecli_version: str,
) -> dict[str, Any]:
    image_path = asset_dir / f"{row['case_uid']}.png"
    render_unit(row, image_path)
    rubric = RUBRIC_BY_FORMAT[row["format"]]
    outputs = {
        dimension: score_dimension(
            client=client,
            model=model,
            image_path=image_path,
            case_uid=row["case_uid"],
            format_name=row["format"],
            dimension=dimension,
            reasoning_effort=reasoning_effort,
            timeout_s=timeout_s,
            max_retries=max_retries,
        )
        for dimension in ("aesthetics", "communication_effectiveness")
    }
    results = []
    for dimension in ("aesthetics", "content_accuracy", "communication_effectiveness"):
        definitions = flatten(rubric, dimension)
        if dimension == "content_accuracy":
            for definition in definitions:
                results.append({
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
                })
            continue
        scored = {item["subquestion_id"]: item for item in outputs[dimension]["results"]}
        for definition in definitions:
            assessment = scored[definition["id"]]
            results.append({
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
            })
    return {
        "case_uid": row["case_uid"],
        "format": row["format"],
        "unit_type": row["unit_type"],
        "unit_name": row["unit_name"],
        "slide_id": row["case_uid"],
        "case_id": row["case_uid"],
        "title": Path(row["source_name"]).stem.replace("_", " ").replace("-", " "),
        "role": row["unit_name"],
        "image": f"/benchmark-units-v3/{image_path.name}",
        "image_sha256": sha256_bytes(image_path.read_bytes()),
        "human_aesthetics_grade_0_3": None,
        "human_agreement": None,
        "human_reason": "",
        "sample_source": row["source_set"],
        "source_name": row["source_name"],
        "source_document_sha256": row["source_document_sha256"],
        "evidence_source": "v3_new_model_calls",
        "render_provenance": {"officecli_version": officecli_version},
        "results": results,
        "criterion_rollups": criterion_rollups(rubric, results),
        "calls": {
            dimension: {key: value for key, value in output.items() if key != "results"}
            for dimension, output in outputs.items()
        },
    }


def prepare_ppt_cases(v2: dict[str, Any]) -> list[dict[str, Any]]:
    cases = []
    for item in v2["cases"]:
        case = copy.deepcopy(item)
        case.update({
            "case_uid": f"pptx-{item['slide_id']}",
            "format": "pptx",
            "unit_type": "slide",
            "unit_name": item["role"],
            "source_name": None,
            "source_document_sha256": None,
            "evidence_source": "v2_reused",
            "render_provenance": {"source": "PPT V2 tracked benchmark render"},
        })
        cases.append(case)
    return cases


def summarize(cases: list[dict[str, Any]], v2: dict[str, Any]) -> dict[str, Any]:
    new_cases = [case for case in cases if case["evidence_source"] == "v3_new_model_calls"]
    new_tokens = sum(
        int((call.get("usage") or {}).get("total_tokens") or 0)
        for case in new_cases
        for call in case["calls"].values()
    )
    return {
        "format_counts": dict(Counter(case["format"] for case in cases)),
        "scored_subquestion_count": sum(
            row["status"] == "scored" for case in cases for row in case["results"]
        ),
        "abstention_count": sum(
            row["status"] == "not_assessable" for case in cases for row in case["results"]
        ),
        "new_total_tokens": new_tokens,
        "represented_total_tokens": new_tokens + int(v2["summary"]["total_tokens"]),
        "ppt_human_ai_aesthetics_spearman": v2["summary"]["human_ai_aesthetics_spearman"],
        "ppt_human_ai_pairwise_accuracy": v2["summary"]["human_ai_pairwise_accuracy"],
        "human_labels": {"pptx": 30, "docx": 0, "xlsx": 0},
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--v2", type=Path, default=DEFAULT_V2)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--asset-dir", type=Path, default=DEFAULT_ASSET_DIR)
    parser.add_argument("--run-dir", type=Path, default=DEFAULT_RUN_DIR)
    parser.add_argument("--corpus-root", type=Path, default=DEFAULT_CORPUS_ROOT)
    parser.add_argument("--base-url", default="http://127.0.0.1:19100/v1")
    parser.add_argument("--model", default="gpt-5.5")
    parser.add_argument("--reasoning-effort", default="medium")
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--timeout", type=float, default=300)
    parser.add_argument("--max-retries", type=int, default=2)
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()

    v2 = json.loads(args.v2.read_text(encoding="utf-8"))
    docs = document_manifest(args.corpus_root)
    args.run_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = args.run_dir / "manifest.json"
    public_docs = [
        {key: value for key, value in row.items() if not key.startswith("_")}
        for row in docs
    ]
    atomic_json(manifest_path, public_docs)
    officecli_version = subprocess.check_output(
        ["officecli", "--version"], text=True, timeout=20
    ).strip()
    rubric_hash = sha256_bytes(canonical_json(RUBRIC_BY_FORMAT).encode("utf-8"))
    progress_metadata = {
        "model": args.model,
        "reasoning_effort": args.reasoning_effort,
        "rubric_sha256": rubric_hash,
        "manifest_sha256": sha256_bytes(manifest_path.read_bytes()),
        "v2_sha256": sha256_bytes(args.v2.read_bytes()),
        "officecli_version": officecli_version,
    }
    progress_path = args.run_dir / "progress.json"
    progress = {"metadata": progress_metadata, "cases": {}}
    if progress_path.exists() and not args.force:
        progress = json.loads(progress_path.read_text(encoding="utf-8"))
        if progress.get("metadata") != progress_metadata:
            raise ValueError("progress metadata does not match this V3 run")
    completed = dict(progress.get("cases") or {})
    client = OpenAI(base_url=args.base_url.rstrip("/"), api_key="local-proxy", max_retries=0)
    lock = threading.Lock()

    def run(row: dict[str, Any]) -> tuple[str, dict[str, Any]]:
        key = row["case_uid"]
        if key in completed and not args.force:
            return key, completed[key]
        result = score_document_case(
            row,
            client=client,
            model=args.model,
            reasoning_effort=args.reasoning_effort,
            timeout_s=args.timeout,
            max_retries=args.max_retries,
            asset_dir=args.asset_dir,
            officecli_version=officecli_version,
        )
        with lock:
            completed[key] = result
            atomic_json(progress_path, {"metadata": progress_metadata, "cases": completed})
        return key, result

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        futures = {pool.submit(run, row): row["case_uid"] for row in docs}
        for future in as_completed(futures):
            key, _ = future.result()
            print(f"scored {key} ({len(completed)}/24)", flush=True)
    atomic_json(progress_path, {"metadata": progress_metadata, "cases": completed})

    cases = prepare_ppt_cases(v2) + [completed[row["case_uid"]] for row in docs]
    artifact = {
        "experiment": {
            "id": "office-subquestion-v3",
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "model": args.model,
            "provider": "local-responses-compatible",
            "reasoning_effort": args.reasoning_effort,
            "image_detail": "high",
            "rubric_version": RUBRIC_VERSION,
            "rubric_sha256": rubric_hash,
            "case_count": 54,
            "reused_case_count": 30,
            "new_case_count": 24,
            "new_model_call_count": 48,
            "represented_model_call_count": 108,
            "subquestions_per_criterion": 4,
            "content_accuracy_policy": ABSTENTION_REASON,
            "officecli_version": officecli_version,
        },
        "summary": summarize(cases, v2),
        "rubric_by_format": RUBRIC_BY_FORMAT,
        "cases": cases,
    }
    atomic_json(args.output, artifact)
    print(f"wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
