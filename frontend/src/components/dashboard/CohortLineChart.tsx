import {
  ResponsiveContainer,
  LineChart,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
} from "recharts";
import type { CohortData, CohortEntry, CohortMetric } from "../../types/dashboard";
import { formatCurrency } from "../../utils/format";

interface Props {
  cohort: CohortData;
  metric: CohortMetric;
  scaleFactor: number;
  granularity: string;
  metricLabel?: string;
}

const COLORS = [
  "#14B8A6", "#F59E0B", "#6366F1", "#EC4899", "#10B981",
  "#8B5CF6", "#F97316", "#06B6D4", "#EF4444", "#84CC16",
  "#A855F7", "#0EA5E9", "#F43F5E", "#22D3EE", "#D946EF",
];

const PERIOD_PREFIX: Record<string, string> = {
  annual: "Y",
  quarterly: "Q",
  monthly: "M",
};

function getMetricValues(c: CohortEntry, metric: CohortMetric): (number | null)[] {
  switch (metric) {
    case "ndr": return c.ndr;
    case "logo_retention": return c.logo_retention;
    case "arr": return c.arr;
    case "customers": return c.customers;
  }
}

function leftAlign(values: (number | null)[]): (number | null)[] {
  const first = values.findIndex((v) => v != null);
  return first >= 0 ? values.slice(first) : [];
}

/** Produce evenly-spaced round tick values that cover [min, max]. */
function niceTicks(min: number, max: number, targetCount = 5): number[] {
  const range = max - min || max * 0.1 || 1;
  const rough = range / targetCount;
  const pow = Math.pow(10, Math.floor(Math.log10(rough)));
  const norm = rough / pow;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10) * pow;
  const lo = Math.floor(min / step) * step;
  const hi = Math.ceil(max / step) * step;
  const ticks: number[] = [];
  for (let v = lo; v <= hi + step * 1e-9; v += step) {
    ticks.push(Math.round(v * 1e10) / 1e10);
  }
  return ticks;
}

export default function CohortLineChart({ cohort, metric, scaleFactor, granularity }: Props) {
  const { cohorts } = cohort;
  if (!cohorts?.length) return null;

  const aligned = cohorts.map((c) => ({
    c,
    shifted: leftAlign(getMetricValues(c, metric)),
  }));

  const maxCols = Math.max(...aligned.map((a) => a.shifted.length));
  const prefix = PERIOD_PREFIX[granularity] ?? "Y";

  const chartData = Array.from({ length: maxCols }, (_, colIdx) => {
    const point: Record<string, number | null | string> = {
      period: `${prefix}${colIdx}`,
    };
    for (const { c, shifted } of aligned) {
      point[c.label] = shifted[colIdx] ?? null;
    }
    return point;
  });

  const isPct = metric === "ndr" || metric === "logo_retention";
  const isCurrency = metric === "arr";

  // Compute nice round Y-axis ticks from data range
  const allValues: number[] = [];
  for (const { shifted } of aligned) {
    for (const v of shifted) {
      if (v != null) allValues.push(v);
    }
  }
  const dataMin = Math.min(...allValues);
  const dataMax = Math.max(...allValues);
  const yTicks = niceTicks(dataMin, dataMax);
  const yDomain: [number, number] = [yTicks[0], yTicks[yTicks.length - 1]];

  const yFormatter = (v: number) => {
    if (isPct) return `${Math.round(v * 100)}%`;
    if (isCurrency) return formatCurrency(v, scaleFactor);
    return v.toLocaleString();
  };

  const tooltipFormatter = (value: number) => {
    if (isPct) return `${Math.round(value * 100)}%`;
    if (isCurrency) return formatCurrency(value, scaleFactor);
    return value.toLocaleString();
  };

  // Map cohort label → index for chronological tooltip sorting
  const cohortOrder = new Map(aligned.map(({ c }, i) => [c.label, i]));

  return (
    <div className="max-w-3xl">
      {/* Legend */}
      <div className="flex flex-wrap gap-x-4 gap-y-1 mb-2">
        {aligned.map(({ c }, i) => (
          <div key={c.label} className="flex items-center gap-1.5">
            <div
              className="w-3 h-0.5 rounded-full"
              style={{ backgroundColor: COLORS[i % COLORS.length] }}
            />
            <span className="text-xs text-gray-500">{c.label}</span>
          </div>
        ))}
      </div>

      <ResponsiveContainer width="100%" height={220}>
        <LineChart data={chartData} margin={{ top: 5, right: 20, bottom: 5, left: 10 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#E5E7EB" opacity={0.8} vertical={false} />
          <XAxis
            dataKey="period"
            tick={{ fontSize: 11, fill: "#6B7280" }}
            axisLine={{ stroke: "#D1D5DB" }}
            tickLine={false}
          />
          <YAxis
            domain={yDomain}
            ticks={yTicks}
            tickFormatter={yFormatter}
            tick={{ fontSize: 11, fill: "#6B7280" }}
            axisLine={false}
            tickLine={false}
          />
          <Tooltip
            formatter={tooltipFormatter}
            itemSorter={(item: any) => cohortOrder.get(item.dataKey) ?? 0}
            contentStyle={{
              backgroundColor: "#ffffff",
              border: "1px solid #E5E7EB",
              borderRadius: "8px",
              color: "#111827",
              fontSize: 12,
            }}
          />
          {aligned.map(({ c }, i) => (
            <Line
              key={c.label}
              type="monotone"
              dataKey={c.label}
              stroke={COLORS[i % COLORS.length]}
              strokeWidth={1.5}
              dot={false}
              connectNulls={false}
              isAnimationActive={false}
            />
          ))}
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}
