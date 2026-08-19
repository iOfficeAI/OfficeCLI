import type { OfficeFormat, RewardCase } from "../case-data";

export type DimensionKey = keyof RewardCase["scores"];

export type RubricCriterion = {
  id: string;
  dimension: DimensionKey;
  label: string;
  prompt: string;
  five: string;
  four: string;
  bias: number;
};

type CriterionSeed = Omit<RubricCriterion, "dimension">;

const dimensions: Record<
  DimensionKey,
  { label: string; shortLabel: string; weight: number }
> = {
  aesthetics: {
    label: "Aesthetics",
    shortLabel: "美观",
    weight: 0.4,
  },
  content_accuracy: {
    label: "Content Accuracy",
    shortLabel: "准确",
    weight: 0.35,
  },
  communication_effectiveness: {
    label: "Communication Effectiveness",
    shortLabel: "传达",
    weight: 0.25,
  },
};

function dimensionCriteria(
  dimension: DimensionKey,
  items: CriterionSeed[],
): RubricCriterion[] {
  return items.map((item) => ({ ...item, dimension }));
}

const pptCriteria: RubricCriterion[] = [
  ...dimensionCriteria("aesthetics", [
    {
      id: "ppt-layout",
      label: "布局与构图",
      prompt: "检查画面平衡、对齐、边距、分组，以及 16:9 画布是否被有意使用。",
      five: "构图平衡、对齐精准、留白有目的，视觉重心明确且没有未完成区域",
      four: "结构和分组清楚，仅有轻微间距或比例不均，不妨碍整体专业度",
      bias: 0.2,
    },
    {
      id: "ppt-typography",
      label: "字体与层级",
      prompt: "检查演示距离下的字号、行长、对比度、标题层级和字体一致性。",
      five: "演示距离下完全易读，标题与正文层级稳定，字号和字重使用高度一致",
      four: "整体易读且层级明确，但存在一处偏小文字、行长或样式不一致",
      bias: 0,
    },
    {
      id: "ppt-graphics",
      label: "图表与图形",
      prompt: "检查图表、图片、表格、图标、箭头和连接线的质量与整合程度。",
      five: "图形清晰、标注完整、风格一致，并直接支持页面的主要信息",
      four: "图形可读且与内容相关，仅有轻微标注、对齐或样式问题",
      bias: 0.1,
    },
    {
      id: "ppt-technical",
      label: "技术完整性",
      prompt: "优先检查重叠、裁切、截断、溢出、断裂连接线和未完成区域。",
      five: "没有可见重叠、裁切、溢出或断裂，所有元素均完整呈现",
      four: "没有实质性故障，仅有不影响阅读的轻微边界或渲染问题",
      bias: 0.3,
    },
    {
      id: "ppt-economy",
      label: "视觉经济性",
      prompt: "检查是否同时避免无理由的大块空白和文档式拥挤。",
      five: "信息密度恰当，留白帮助阅读，没有冗余装饰或拥挤区域",
      four: "密度整体合理，但有一处略空、略挤或可精简的内容",
      bias: -0.1,
    },
  ]),
  ...dimensionCriteria("content_accuracy", [
    {
      id: "ppt-central-claim",
      label: "中心结论",
      prompt: "识别页面中心结论，并与提供的 reference evidence 逐项核对。",
      five: "中心结论被参考证据直接支持，方向、范围和措辞均准确",
      four: "中心结论正确，仅有不改变结论的轻微限定或措辞偏差",
      bias: 0.2,
    },
    {
      id: "ppt-numbers",
      label: "数字与计算",
      prompt: "核对数字、百分比、公式、表格值和图表数据。",
      five: "所有可核验数字、单位、计算和图表值与参考证据精确一致",
      four: "关键数字正确，仅有轻微舍入、精度或格式差异",
      bias: 0.1,
    },
    {
      id: "ppt-labels",
      label: "标签与来源",
      prompt: "核对方法名、数据集、基线、图例、坐标轴和归因标签。",
      five: "方法、数据集、基线和标签完整准确，来源归因无歧义",
      four: "主要标签准确，仅有一处次要标签、缩写或来源说明不足",
      bias: 0,
    },
    {
      id: "ppt-scope",
      label: "范围与因果",
      prompt: "检查比较、因果、最高级、全称和实时性等范围声明。",
      five: "比较范围、因果强度和限定条件均与证据严格一致",
      four: "总体范围正确，但有一处限定条件可写得更精确",
      bias: -0.1,
    },
    {
      id: "ppt-reference-coverage",
      label: "证据覆盖",
      prompt: "判断重要可见声明是否都能由 reference evidence 覆盖并验证。",
      five: "每个重要声明都能被直接核验，没有遗漏或不可验证的核心内容",
      four: "核心内容均可验证，仅有一个非关键细节缺少直接证据",
      bias: 0.2,
    },
  ]),
  ...dimensionCriteria("communication_effectiveness", [
    {
      id: "ppt-takeaway",
      label: "核心信息清晰度",
      prompt: "仅根据可见内容，用一句话复述 intended takeaway，并判断是否稳定。",
      five: "数秒内即可准确复述唯一核心信息，不需要猜测或补充上下文",
      four: "核心信息清楚，但需要一次短暂阅读或整合两个相邻信息块",
      bias: 0.2,
    },
    {
      id: "ppt-alignment",
      label: "标题与证据一致性",
      prompt: "检查标题、强调、证据和结论是否共同支持同一信息。",
      five: "标题、视觉强调、证据和结论完全同向，没有竞争性信息",
      four: "主线一致，仅有一个次要元素的强调程度略有偏差",
      bias: 0.1,
    },
    {
      id: "ppt-reading-path",
      label: "阅读路径",
      prompt: "判断第一、第二和最终视觉落点是否形成连贯顺序。",
      five: "阅读顺序自然且唯一，分组、对齐和连接关系无需解释",
      four: "主要路径连贯，仅有一次轻微跳转或局部顺序可优化",
      bias: 0,
    },
    {
      id: "ppt-comprehension",
      label: "受众理解成本",
      prompt: "检查术语、缩写、图例、密度和标签是否适合目标受众。",
      five: "目标受众无需反复阅读即可理解，术语和图例解释充分",
      four: "整体容易理解，但有一个缩写、标签或密集区域增加阅读成本",
      bias: -0.1,
    },
    {
      id: "ppt-information-economy",
      label: "信息取舍",
      prompt: "检查每个内容块是否直接支持 takeaway，是否存在重复或应移入备注的细节。",
      five: "每个元素都服务于结论，信息精炼且没有可见重复",
      four: "绝大多数内容必要，仅有一小块信息可删减或移入备注",
      bias: 0.1,
    },
  ]),
];

const wordCriteria: RubricCriterion[] = [
  ...dimensionCriteria("aesthetics", [
    {
      id: "word-page-composition",
      label: "页面构图与页边距",
      prompt: "检查可用正文区域、边距、对齐、分页和平衡是否跨页面一致。",
      five: "页面平衡、边距统一、分页自然，正文区域使用充分且没有突兀空白",
      four: "整体出版级，仅有一处分页、页边距或页面平衡的小问题",
      bias: 0.1,
    },
    {
      id: "word-typography",
      label: "字体与标题层级",
      prompt: "检查正文可读性、样式一致性、标题级别、强调和章节区分。",
      five: "正文舒适易读，标题层级严格一致，样式能稳定表达文档结构",
      four: "层级和正文均清楚，仅有一处字号、字重或样式使用不一致",
      bias: 0.1,
    },
    {
      id: "word-spacing",
      label: "间距与阅读节奏",
      prompt: "检查段落、列表、标题、表格、图表和图片之间的间距。",
      five: "垂直节奏稳定，段落和对象间距准确帮助连续阅读",
      four: "整体节奏清晰，但有一处略紧、略松或分页附近的间距不佳",
      bias: 0,
    },
    {
      id: "word-media",
      label: "表格图表与图片",
      prompt: "检查大小、位置、题注、环绕、对齐、可读性及与正文的整合。",
      five: "所有对象清晰、题注完整、位置稳定，并与引用它们的正文紧密对应",
      four: "对象可读且整合良好，仅有轻微尺寸、题注或环绕问题",
      bias: -0.1,
    },
    {
      id: "word-running-elements",
      label: "页眉页脚",
      prompt: "检查页眉、页脚、页码的一致性、克制度和与正文的分隔。",
      five: "页眉页脚一致、克制、页码明确，完全不干扰正文",
      four: "功能完整且一致，仅有一处间距、分隔或样式可优化",
      bias: 0.2,
    },
    {
      id: "word-technical",
      label: "技术完整性",
      prompt: "检查裁切、溢出、孤立标题、落单题注、拥挤页边和未完成格式。",
      five: "没有裁切、溢出、孤立标题、落单题注或未完成格式",
      four: "无内容损失，仅有一处低严重度分页或排版提示",
      bias: 0.3,
    },
  ]),
  ...dimensionCriteria("content_accuracy", [
    {
      id: "word-claims",
      label: "正文与标题声明",
      prompt: "核对正文、标题和摘要中的重要事实与结论。",
      five: "重要声明与参考材料逐项一致，没有夸大、遗漏或方向错误",
      four: "核心声明准确，仅有不改变含义的轻微措辞或限定偏差",
      bias: 0.2,
    },
    {
      id: "word-values",
      label: "数值、单位与日期",
      prompt: "核对数值、单位、日期、人名、机构名和专有名词。",
      five: "所有可核验数值、单位、日期和名称均精确一致",
      four: "关键内容正确，仅有轻微舍入、格式或非关键名称差异",
      bias: 0.1,
    },
    {
      id: "word-cross-references",
      label: "交叉引用与题注",
      prompt: "核对图表编号、题注、章节引用和正文中的交叉引用。",
      five: "所有编号、题注和交叉引用均指向正确对象且内容一致",
      four: "引用关系正确，仅有一处格式、编号样式或说明不够完整",
      bias: 0,
    },
    {
      id: "word-citations",
      label: "引用与注释",
      prompt: "核对引用、脚注、尾注和来源归因是否支持相应声明。",
      five: "每项重要外部声明都有准确、可定位且格式一致的来源",
      four: "主要来源完整，仅有一个次要引用的格式或定位信息不足",
      bias: -0.1,
    },
    {
      id: "word-conclusions",
      label: "结论与证据对应",
      prompt: "检查结论和建议是否由正文证据直接支持，并区分矛盾与证据未覆盖。",
      five: "结论严格来自文中证据，证据覆盖范围和可验证性表达准确",
      four: "结论总体受支持，但有一个次要推断需要更明确的限定",
      bias: 0.2,
    },
  ]),
  ...dimensionCriteria("communication_effectiveness", [
    {
      id: "word-purpose",
      label: "目的清晰度",
      prompt: "判断文档单元是否明确说明目标、受众和预期结果。",
      five: "开头即可明确识别目的、受众和预期行动",
      four: "目的清楚，但受众或预期结果需要从上下文进一步确认",
      bias: 0.2,
    },
    {
      id: "word-progression",
      label: "章节推进",
      prompt: "检查标题和段落是否按合理顺序推进论点或工作流。",
      five: "章节顺序自然，每一部分都承接前文并推进结论",
      four: "主线连贯，仅有一个段落或小节的位置可调整",
      bias: 0.1,
    },
    {
      id: "word-navigation",
      label: "导航能力",
      prompt: "检查标题、列表、题注、表格、交叉引用、页眉页脚是否帮助定位。",
      five: "读者可快速定位任意信息，层级与引用系统完整一致",
      four: "整体易导航，但有一处标题层级、列表或交叉引用提示不足",
      bias: 0,
    },
    {
      id: "word-density",
      label: "冗余与密度",
      prompt: "检查重复声明、可避免的长句、拥挤页面和中断理解的稀疏片段。",
      five: "内容紧凑但不拥挤，没有重复，段落长度和页面密度恰当",
      four: "整体节制，仅有一段偏密、偏长或可删除的重复信息",
      bias: -0.1,
    },
    {
      id: "word-audience-detail",
      label: "受众适配",
      prompt: "检查定义、证据、技术深度和摘要是否适合目标读者。",
      five: "术语、背景、证据深度和总结完全匹配目标读者",
      four: "整体匹配，仅有一个术语或技术细节需要补充说明或精简",
      bias: 0.1,
    },
  ]),
];

const excelCriteria: RubricCriterion[] = [
  ...dimensionCriteria("aesthetics", [
    {
      id: "excel-used-range",
      label: "有效区域布局",
      prompt: "检查 populated regions 是否紧凑、可读、对齐且形成有意结构。",
      five: "有效区域紧凑、对齐稳定、区块关系清楚，没有散落内容",
      four: "整体结构明确，仅有一处间距、对齐或区块位置可优化",
      bias: 0.2,
    },
    {
      id: "excel-sizing",
      label: "列宽与行高",
      prompt: "检查截断、过度换行、尺寸不一致和无效空间。",
      five: "所有内容完整显示，列宽行高一致且适配信息密度",
      four: "整体可读，仅有一个次要列或行略宽、略窄或发生非关键换行",
      bias: 0,
    },
    {
      id: "excel-number-format",
      label: "数字格式",
      prompt: "检查日期、百分比、货币、小数、正负号、单位和精度的一致性。",
      five: "数字格式完全一致，单位和精度明确，比较无需额外解释",
      four: "主要格式一致，仅有一处非关键精度、符号或单位展示问题",
      bias: 0.1,
    },
    {
      id: "excel-hierarchy",
      label: "输入计算输出层级",
      prompt: "检查可编辑假设、计算过程和结果是否通过样式与位置清楚区分。",
      five: "输入、计算和输出一眼可辨，样式语义稳定且没有混淆",
      four: "层级总体清楚，但有一个区域的输入或结果强调不足",
      bias: 0.2,
    },
    {
      id: "excel-charts",
      label: "图表与标签",
      prompt: "检查图表尺寸、位置、标题、坐标轴、图例、数据标签和源上下文。",
      five: "图表完整易读、标签齐全、位置合理，并与源数据关系明确",
      four: "图表可用且结论清楚，仅有一处标签、图例或位置问题",
      bias: 0,
    },
    {
      id: "excel-conditional-format",
      label: "条件格式",
      prompt: "检查条件格式是否克制、可解释、有用且不会误导。",
      five: "条件格式规则清楚、颜色克制，并准确突出需要行动的异常",
      four: "整体有效，仅有一个颜色、阈值说明或覆盖范围可优化",
      bias: -0.1,
    },
    {
      id: "excel-scanability",
      label: "留白与扫描效率",
      prompt: "检查过量空白、散落内容、焦点顺序和 dashboard 扫描效率。",
      five: "首屏即可完成从摘要到细节的扫描，留白恰当且无分散区域",
      four: "整体扫描顺畅，仅有一处空白偏多或次要区域略分散",
      bias: 0.1,
    },
  ]),
  ...dimensionCriteria("content_accuracy", [
    {
      id: "excel-values",
      label: "值、单位与汇总",
      prompt: "核对可见值、单位、日期、百分比、合计、小计和输出。",
      five: "所有关键值、单位、合计和小计与参考数据精确一致",
      four: "主要结果正确，仅有轻微舍入、格式或非关键汇总差异",
      bias: 0.2,
    },
    {
      id: "excel-formulas",
      label: "公式与计算结果",
      prompt: "检查公式逻辑，并在可用时核对 cached values 与 computed results。",
      five: "公式逻辑正确，缓存值与计算结果一致，没有错误或陈旧输出",
      four: "核心公式正确，仅有一个非关键公式可简化或存在轻微缓存提示",
      bias: 0.1,
    },
    {
      id: "excel-chart-sources",
      label: "图表数据源",
      prompt: "核对图表源范围、series、category labels 和显示结果。",
      five: "所有图表范围、系列和分类标签均指向正确数据且无遗漏",
      four: "主要图表数据正确，仅有一个次要标签或范围边界可改进",
      bias: 0,
    },
    {
      id: "excel-references",
      label: "命名与跨表引用",
      prompt: "检查 named ranges、cross-sheet references 及隐藏或缺失单元格引用。",
      five: "命名范围和跨表引用全部有效、可追踪且没有隐藏依赖风险",
      four: "引用正确，仅有一个命名、可读性或隐藏依赖说明不足",
      bias: -0.1,
    },
    {
      id: "excel-assumptions",
      label: "假设与错误类型",
      prompt: "区分公式错误、引用错误、单位不匹配、业务假设和不可验证内容。",
      five: "假设明确、单位一致，错误类型与不可验证项均被准确标识",
      four: "核心假设合理且可追踪，仅有一个次要假设缺少明确说明",
      bias: 0.1,
    },
  ]),
  ...dimensionCriteria("communication_effectiveness", [
    {
      id: "excel-workflow",
      label: "输入到输出工作流",
      prompt: "判断读者能否识别输入、计算、输出、趋势、异常和决策。",
      five: "无需反向推导即可理解输入到输出的完整工作流",
      four: "主要工作流清楚，但一个中间计算或状态需要短暂查找",
      bias: 0.2,
    },
    {
      id: "excel-labeling",
      label: "标签、单位与图例",
      prompt: "检查标签、单位、图例、说明和分组是否让内容自解释。",
      five: "所有区域、数值、图表和状态都有明确标签与单位",
      four: "大部分内容自解释，仅有一个次要单位、图例或说明缺失",
      bias: 0.1,
    },
    {
      id: "excel-context",
      label: "层级与工作表上下文",
      prompt: "检查 freeze panes、视觉层级、sheet context 和摘要区域是否支持定位。",
      five: "层级和工作表上下文完整，摘要与明细之间可快速定位",
      four: "上下文总体明确，但有一个冻结、导航或摘要关联可优化",
      bias: 0,
    },
    {
      id: "excel-explainability",
      label: "假设与公式可解释性",
      prompt: "检查假设、关键公式和例外是否有足够说明。",
      five: "关键假设、公式和例外均有清楚说明，不需要猜测",
      four: "核心逻辑可理解，仅有一个次要假设或公式缺少注释",
      bias: -0.1,
    },
    {
      id: "excel-decision",
      label: "决策相关性",
      prompt: "检查输出、趋势、异常和结论是否明确指向需要采取的决策。",
      five: "关键结果和异常直接对应明确决策，优先级一目了然",
      four: "决策方向清楚，但一个次要异常或行动优先级可更明确",
      bias: 0.1,
    },
  ]),
];

export const criteriaByFormat: Record<OfficeFormat, RubricCriterion[]> = {
  pptx: pptCriteria,
  docx: wordCriteria,
  xlsx: excelCriteria,
};

export const formatLabels: Record<
  OfficeFormat,
  { short: string; long: string }
> = {
  pptx: { short: "PPT", long: "PowerPoint" },
  docx: { short: "Word", long: "Word" },
  xlsx: { short: "Excel", long: "Excel" },
};

export const dimensionLabels = dimensions;

export const anchorCases: Record<
  OfficeFormat,
  Record<
    DimensionKey,
    {
      five: { title: string; evidence: string };
      four: { title: string; evidence: string };
    }
  >
> = {
  pptx: {
    aesthetics: {
      five: {
        title: "Product Roadmap",
        evidence: "里程碑序列、owner 与 quarter 标签清晰，没有重叠。",
      },
      four: {
        title: "Research Results",
        evidence: "整体专业，但存在两处字体替换和轻微密度问题。",
      },
    },
    content_accuracy: {
      five: {
        title: "Executive Review",
        evidence: "KPI 标签和数值均与参考证据一致。",
      },
      four: {
        title: "Product Roadmap",
        evidence: "核心里程碑准确，部分计划性陈述仍依赖未来验证。",
      },
    },
    communication_effectiveness: {
      five: {
        title: "Product Roadmap",
        evidence: "每页均可直接恢复 foundation 到 scale 的主线。",
      },
      four: {
        title: "Research Results",
        evidence: "实验与 ablation 顺序清楚，但信息密度略高。",
      },
    },
  },
  docx: {
    aesthetics: {
      five: {
        title: "Policy Brief",
        evidence: "编号、标题层级和短页结构稳定，未发现渲染问题。",
      },
      four: {
        title: "Table Analysis",
        evidence: "表格完整但密度偏高，正常阅读需要更慢扫描。",
      },
    },
    content_accuracy: {
      five: {
        title: "Policy Brief",
        evidence: "政策声明、编号建议和参考证据完全对应。",
      },
      four: {
        title: "Table Analysis",
        evidence: "比较值准确，复杂表格中的次要说明仍可加强。",
      },
    },
    communication_effectiveness: {
      five: {
        title: "Policy Brief",
        evidence: "标题层级和编号建议让读者快速定位行动项。",
      },
      four: {
        title: "Table Analysis",
        evidence: "结构可恢复，但表格密度增加了理解成本。",
      },
    },
  },
  xlsx: {
    aesthetics: {
      five: {
        title: "Chart Dashboard",
        evidence: "KPI tile 与图表扫描顺序紧凑明确，chart crop 完整。",
      },
      four: {
        title: "Operations Sheet",
        evidence: "状态格式清楚，但存在一个数字格式提示。",
      },
    },
    content_accuracy: {
      five: {
        title: "Revenue Model",
        evidence: "公式、缓存值、计算结果和图表引用全部一致。",
      },
      four: {
        title: "Chart Dashboard",
        evidence: "KPI 与源单元格一致，复杂聚合仍需要保留来源追踪。",
      },
    },
    communication_effectiveness: {
      five: {
        title: "Chart Dashboard",
        evidence: "摘要、趋势和 KPI 的扫描路径直接指向决策。",
      },
      four: {
        title: "Revenue Model",
        evidence: "两张表工作流清楚，但模型计算区需要一定追踪。",
      },
    },
  },
};

export function subscorePrompt(criterion: RubricCriterion): string {
  return [
    `只评估“${criterion.label}”。`,
    criterion.prompt,
    "按 1–5 的整数评分；先列出可见证据，再给分。",
    "不得用其他小项的优点补偿本项缺陷；不确定时取较低分。",
  ].join(" ");
}
