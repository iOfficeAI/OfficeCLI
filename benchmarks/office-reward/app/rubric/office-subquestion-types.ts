import rawOfficeSubquestionResults from "./office-subquestion-results-v3.json";

export type OfficeFormat = "pptx" | "docx" | "xlsx";
export type DimensionKey =
  | "aesthetics"
  | "content_accuracy"
  | "communication_effectiveness";

export type SubquestionDefinition = {
  id: string;
  label: string;
  question: string;
};

export type CriterionDefinition = {
  id: string;
  label: string;
  subquestions: SubquestionDefinition[];
};

export type SubquestionResult = {
  criterion_id: string;
  criterion_label: string;
  subquestion_id: string;
  subquestion_label: string;
  question: string;
  dimension: DimensionKey;
  status: "scored" | "not_assessable";
  source: "model" | "abstention";
  score_1_5: number | null;
  confidence_0_1: number;
  evidence: string;
  defects: string[];
};

export type CriterionRollup = {
  criterion_id: string;
  criterion_label: string;
  dimension: DimensionKey;
  score_1_5: number | null;
  source: "transparent_mean_of_direct_subquestions" | "abstention";
};

export type OfficeSubquestionCase = {
  case_uid: string;
  format: OfficeFormat;
  unit_type: "slide" | "page" | "sheet";
  unit_name: string;
  slide_id: string;
  case_id: string;
  title: string;
  role: string;
  image: string;
  image_sha256: string;
  human_aesthetics_grade_0_3: number | null;
  human_agreement: number | null;
  human_reason: string;
  sample_source: string;
  source_name: string | null;
  source_document_sha256: string | null;
  evidence_source: "v2_reused" | "v3_new_model_calls";
  results: SubquestionResult[];
  criterion_rollups: CriterionRollup[];
  calls: Record<string, { usage: Record<string, number> | null }>;
};

export type OfficeSubquestionArtifact = {
  experiment: {
    id: string;
    generated_at: string;
    model: string;
    reasoning_effort: string;
    rubric_sha256: string;
    case_count: number;
    represented_model_call_count: number;
    new_model_call_count: number;
    officecli_version: string;
  };
  summary: {
    format_counts: Record<OfficeFormat, number>;
    scored_subquestion_count: number;
    abstention_count: number;
    new_total_tokens: number;
    represented_total_tokens: number;
    ppt_human_ai_aesthetics_spearman: number;
    ppt_human_ai_pairwise_accuracy: number;
    human_labels: Record<OfficeFormat, number>;
  };
  rubric_by_format: Record<
    OfficeFormat,
    Record<DimensionKey, CriterionDefinition[]>
  >;
  cases: OfficeSubquestionCase[];
};

export const officeSubquestionResults =
  rawOfficeSubquestionResults as unknown as OfficeSubquestionArtifact;

export const formatOrder: OfficeFormat[] = ["pptx", "docx", "xlsx"];
export const formatLabels: Record<OfficeFormat, string> = {
  pptx: "PPT",
  docx: "Word",
  xlsx: "Excel",
};

export const dimensionOrder: DimensionKey[] = [
  "aesthetics",
  "content_accuracy",
  "communication_effectiveness",
];
export const dimensionLabels: Record<
  DimensionKey,
  { label: string; short: string }
> = {
  aesthetics: { label: "Aesthetics", short: "美观" },
  content_accuracy: { label: "Content Accuracy", short: "准确" },
  communication_effectiveness: {
    label: "Communication Effectiveness",
    short: "传达",
  },
};
