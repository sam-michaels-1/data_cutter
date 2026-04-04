import type { GridData } from "../../engine/histograms";
import TwoByTwoGrid from "./TwoByTwoGrid";

interface Props {
  netRetention: GridData;
  lossRetention: GridData;
  subtitle?: string;
}

function formatRetention(v: number): string {
  const formatted = `${(Math.abs(v) * 100).toFixed(1)}%`;
  return v < 0 ? `(${formatted})` : formatted;
}

function retentionColor(v: number): { bg: string; text: string } {
  if (v >= 1.5)  return { bg: "rgba(5, 150, 105, 0.38)", text: "#047857" };
  if (v >= 1.2)  return { bg: "rgba(5, 150, 105, 0.28)", text: "#059669" };
  if (v >= 1.1)  return { bg: "rgba(16, 185, 129, 0.22)", text: "#059669" };
  if (v >= 1.0)  return { bg: "rgba(16, 185, 129, 0.14)", text: "#10B981" };
  if (v >= 0.9)  return { bg: "rgba(245, 158, 11, 0.14)", text: "#D97706" };
  if (v >= 0.8)  return { bg: "rgba(249, 115, 22, 0.16)", text: "#EA580C" };
  return { bg: "rgba(239, 68, 68, 0.20)", text: "#DC2626" };
}

export default function RetentionGrids({ netRetention, lossRetention, subtitle }: Props) {
  return (
    <div className="space-y-2">
      <TwoByTwoGrid
        data={netRetention}
        title="Net Retention by Segment"
        subtitle={subtitle}
        formatMetric={formatRetention}
        colorScale={retentionColor}
      />
      <TwoByTwoGrid
        data={lossRetention}
        title="Lost-Only Retention by Segment"
        subtitle={subtitle}
        formatMetric={formatRetention}
        colorScale={retentionColor}
      />
    </div>
  );
}
