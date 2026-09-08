import type { GridData } from "../../engine/histograms";
import TwoByTwoGrid from "./TwoByTwoGrid";
import { retentionColor } from "./colorScales";

interface Props {
  netRetention: GridData;
  lossRetention: GridData;
  subtitle?: string;
}

function formatRetention(v: number): string {
  const formatted = `${(Math.abs(v) * 100).toFixed(1)}%`;
  return v < 0 ? `(${formatted})` : formatted;
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
