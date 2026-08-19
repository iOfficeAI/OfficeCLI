import type { Metadata } from "next";
import RubricWorkbench from "./rubric-workbench";

export const metadata: Metadata = {
  title: "Office 细粒度评分实验",
  description: "54 个 PPT、Word、Excel 单元的板块小分、直接子问题分与人工标注对照。",
};

export default function RubricPage() {
  return <RubricWorkbench />;
}
