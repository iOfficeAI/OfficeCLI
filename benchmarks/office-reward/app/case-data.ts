import rawCases from "./case-data.json";

export type OfficeFormat = "pptx" | "docx" | "xlsx";

export type RewardCase = {
  id: string;
  format: OfficeFormat;
  title: string;
  subtitle: string;
  image: string;
  unitType: "slide" | "page" | "section" | "sheet";
  unitCount: number;
  scores: {
    aesthetics: number;
    content_accuracy: number;
    communication_effectiveness: number;
  };
  overall_raw_score_100: number;
  reward_0_1: number;
  coverage_0_1: 1;
  status: "complete";
  evidence: string[];
  issueSummary: string;
};

export const rewardCases = rawCases as RewardCase[];
