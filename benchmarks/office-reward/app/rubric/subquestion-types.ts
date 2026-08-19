import rawSubquestionResults from "./subquestion-results-v2.json";

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

export type SubquestionCase = {
  slide_id: string;
  case_id: string;
  title: string;
  role: string;
  image: string;
  image_sha256: string;
  human_aesthetics_grade_0_3: number;
  human_agreement: number | null;
  human_reason: string;
  sample_source: string;
  results: SubquestionResult[];
  criterion_rollups: CriterionRollup[];
  calls: Record<
    string,
    {
      response_id: string | null;
      served_model: string;
      usage: Record<string, number> | null;
      prompt_sha256: string;
    }
  >;
};

export type SubquestionArtifact = {
  experiment: {
    id: string;
    generated_at: string;
    model: string;
    provider: string;
    reasoning_effort: string;
    image_detail: string;
    rubric_version: string;
    rubric_sha256: string;
    manifest_sha256: string;
    scoring_mode: string;
    case_count: number;
    model_call_count: number;
    subquestions_per_criterion: number;
    content_accuracy_policy: string;
  };
  summary: {
    scored_subquestion_count: number;
    abstention_count: number;
    total_tokens: number;
    human_grade_distribution: Record<string, number>;
    human_ai_aesthetics_spearman: number;
    human_ai_pairwise_accuracy: number;
    mean_aesthetics_by_human_grade: Record<string, number>;
    scope: string;
  };
  rubric: Record<DimensionKey, CriterionDefinition[]>;
  cases: SubquestionCase[];
};

export const subquestionResults =
  rawSubquestionResults as unknown as SubquestionArtifact;

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
