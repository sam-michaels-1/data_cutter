import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from "recharts";
import type { PieChartData } from "../../engine/histograms";
import { formatCurrency } from "../../utils/format";

interface Props {
  data: PieChartData[];
  scaleFactor: number;
  metricLabel: string;
}

const COLORS = [
  "#14B8A6", "#F59E0B", "#6366F1", "#EC4899", "#10B981",
  "#8B5CF6", "#F97316", "#06B6D4", "#EF4444", "#84CC16",
  "#A855F7", "#0EA5E9", "#F43F5E", "#22D3EE", "#D946EF",
];

/**
 * Build a stable label -> color map shared by both pies in a group, so the same
 * segment (e.g. "Enterprise") always gets the same color even when the ARR pie
 * and the Customer Count pie are sorted differently.
 */
function buildColorMap(
  ...sliceSets: { label: string }[][]
): Map<string, string> {
  const map = new Map<string, string>();
  for (const slices of sliceSets) {
    for (const s of slices) {
      if (!map.has(s.label)) {
        map.set(s.label, COLORS[map.size % COLORS.length]);
      }
    }
  }
  return map;
}

function SinglePie({ slices, title, formatValue, colorMap }: {
  slices: { label: string; value: number; pct: number }[];
  title: string;
  formatValue: (v: number) => string;
  colorMap: Map<string, string>;
}) {
  if (slices.length === 0) return null;

  const colorFor = (label: string, idx: number) =>
    colorMap.get(label) ?? COLORS[idx % COLORS.length];

  return (
    <div>
      <p className="text-xs font-medium text-gray-500 uppercase tracking-wide mb-2 text-center">{title}</p>
      <ResponsiveContainer width="100%" height={180}>
        <PieChart>
          <Pie
            data={slices}
            dataKey="value"
            nameKey="label"
            cx="50%"
            cy="50%"
            outerRadius={65}
            innerRadius={22}
            paddingAngle={1}
          >
            {slices.map((s, idx) => (
              <Cell key={idx} fill={colorFor(s.label, idx)} />
            ))}
          </Pie>
          <Tooltip
            formatter={(value: unknown, name: unknown) => [formatValue(Number(value)), String(name)]}
            contentStyle={{
              backgroundColor: "#fff",
              border: "1px solid #E5E7EB",
              borderRadius: "8px",
              fontSize: 12,
            }}
          />
        </PieChart>
      </ResponsiveContainer>
      {/* Legend */}
      <div className="flex flex-wrap gap-x-3 gap-y-1 justify-center mt-1">
        {slices.slice(0, 8).map((s, i) => (
          <div key={s.label} className="flex items-center gap-1">
            <div className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: colorFor(s.label, i) }} />
            <span className="text-[10px] text-gray-600">{s.label} ({(s.pct * 100).toFixed(1)}%)</span>
          </div>
        ))}
        {slices.length > 8 && (
          <span className="text-[10px] text-gray-400">+{slices.length - 8} more</span>
        )}
      </div>
    </div>
  );
}

export default function IdentifierPieCharts({ data, scaleFactor, metricLabel }: Props) {
  if (data.length === 0) return null;

  return (
    <div className="space-y-2">
      {data.map(({ identifierName, arrSlices, countSlices }) => {
        // Shared across both pies so each segment keeps one color regardless of sort order
        const colorMap = buildColorMap(arrSlices, countSlices);
        return (
          <div key={identifierName} className="bg-white border border-gray-200 rounded-xl p-3">
            <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-2">
              {identifierName} Breakdown
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <SinglePie
                slices={arrSlices}
                title={`${metricLabel} Distribution`}
                formatValue={(v) => formatCurrency(v, scaleFactor)}
                colorMap={colorMap}
              />
              <SinglePie
                slices={countSlices}
                title="Customer Count"
                formatValue={(v) => v.toLocaleString()}
                colorMap={colorMap}
              />
            </div>
          </div>
        );
      })}
    </div>
  );
}
