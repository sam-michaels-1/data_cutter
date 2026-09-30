import type { GridData } from "../../engine/histograms";
import TwoByTwoGrid from "./TwoByTwoGrid";
import { retentionColor } from "./colorScales";

interface Props {
  netRetention: GridData;
  lossRetention: GridData;
  annNetRetention?: GridData;
  annLossRetention?: GridData;
  subtitle?: string;
  annSubtitle?: string;
}

function formatRetention(v: number): string {
  const formatted = `${(Math.abs(v) * 100).toFixed(1)}%`;
  return v < 0 ? `(${formatted})` : formatted;
}

export default function RetentionGrids({ netRetention, lossRetention, annNetRetention, annLossRetention, subtitle, annSubtitle }: Props) {
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
      {annNetRetention && annLossRetention && (
        <>
          <TwoByTwoGrid
            data={annNetRetention}
            title="Annualized Net Retention by Segment"
            subtitle={annSubtitle}
            formatMetric={formatRetention}
            colorScale={retentionColor}
          />
          <TwoByTwoGrid
            data={annLossRetention}
            title="Annualized Lost-Only Retention by Segment"
            subtitle={annSubtitle}
            formatMetric={formatRetention}
            colorScale={retentionColor}
          />
        </>
      )}
    </div>
  );
}
