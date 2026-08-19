import rawFineGrainedResults from "./fine-grained-results.json";

export type DimensionKey =
  | "aesthetics"
  | "content_accuracy"
  | "communication_effectiveness";

export type CriterionDefinition = {
  id: string;
  label: string;
  question: string;
  score_5?: string;
  score_1?: string;
};

export type CriterionResult = {
  criterion_id: string;
  dimension: DimensionKey;
  status: "scored" | "not_assessable" | "failed";
  source: "model" | "abstention";
  score_1_5: number | null;
  confidence_0_1: number;
  evidence: string;
  defects: string[];
};

export type FineGrainedCase = {
  slide_id: string;
  case_id: string;
  title: string;
  role: string;
  image: string;
  image_sha256: string;
  human_aesthetics_grade_0_3: number;
  human_agreement: number | null;
  human_reason: string;
  results: CriterionResult[];
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

export type FineGrainedArtifact = {
  experiment: {
    id: string;
    generated_at: string;
    model: string;
    provider: string;
    reasoning_effort: string;
    image_detail: string;
    rubric_version: string;
    rubric_sha256: string;
    scoring_mode: string;
    case_count: number;
    model_call_count: number;
    content_accuracy_policy: string;
  };
  summary: {
    scored_criterion_count: number;
    abstention_count: number;
    total_tokens: number;
    human_ai_aesthetics_spearman: number;
    mean_aesthetics_by_human_grade: Record<string, number>;
    scope: string;
  };
  criteria: Record<DimensionKey, CriterionDefinition[]>;
  cases: FineGrainedCase[];
};

export const fineGrainedResults =
  rawFineGrainedResults as unknown as FineGrainedArtifact;

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
