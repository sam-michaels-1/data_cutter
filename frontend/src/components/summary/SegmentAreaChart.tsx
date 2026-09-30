import { useState } from "react";
import {
  ResponsiveContainer,
  ComposedChart,
  Area,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  Legend,
} from "recharts";
import type { SummarySection } from "../../engine/summary_compute";
import { formatCurrency } from "../../utils/format";

interface Props {
  dollars: SummarySection;
  pctOfTotal?: SummarySection;
  columns: string[]; // ['All', ...segmentValues]
  scaleFactor: number;
  metricLabel?: string;
}

const COLORS = [
  "#14B8A6", "#F59E0B", "#6366F1", "#EC4899", "#10B981",
  "#8B5CF6", "#F97316", "#06B6D4", "#EF4444", "#84CC16",
  "#A855F7", "#0EA5E9", "#F43F5E", "#22D3EE", "#D946EF",
];

export default function SegmentAreaChart({ dollars, pctOfTotal, columns, scaleFactor, metricLabel = "ARR" }: Props) {
  const [mode, setMode] = useState<"dollars" | "share">("dollars");
  const section = mode === "share" && pctOfTotal ? pctOfTotal : dollars;
  const isShare = section.key === "pct_of_total";
  const segments = columns.slice(1);

  if (!section.rows.length || !segments.length) return null;

  const chartData = section.rows.map(row => {
    const point: Record<string, number | string | null> = { period: row.period };
    row.values.forEach((v, ci) => {
      point[columns[ci]] = v;
    });
    return point;
  });

  const valueFormatter = (v: number) =>
    isShare ? `${(v * 100).toFixed(1)}%` : formatCurrency(v, scaleFactor);

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-3 shrink-0">
      <div className="flex items-center justify-between mb-1">
        <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide">
          {metricLabel} Over Time by Segment
        </h3>
        {pctOfTotal && (
          <div className="flex gap-1 bg-gray-100 rounded-lg p-0.5">
            {(["dollars", "share"] as const).map(m => (
              <button
                key={m}
                onClick={() => setMode(m)}
                className={`px-2 py-0.5 rounded-md text-[11px] font-medium transition ${
                  mode === m ? "bg-teal-600 text-white shadow" : "text-gray-500 hover:text-gray-700"
                }`}
              >
                {m === "dollars" ? "$" : "% Share"}
              </button>
            ))}
          </div>
        )}
      </div>
      <ResponsiveContainer width="100%" height={190}>
        <ComposedChart data={chartData} margin={{ top: 5, right: 20, bottom: 5, left: 10 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#E5E7EB" opacity={0.8} vertical={false} />
          <XAxis
            dataKey="period"
            tick={{ fontSize: 11, fill: "#6B7280" }}
            axisLine={{ stroke: "#D1D5DB" }}
            tickLine={false}
          />
          <YAxis
            tickFormatter={valueFormatter}
            tick={{ fontSize: 11, fill: "#6B7280" }}
            axisLine={false}
            tickLine={false}
            domain={isShare ? [0, 1] : [0, "auto"]}
          />
          <Tooltip
            formatter={(value, name) => [valueFormatter(Number(value)), String(name)]}
            contentStyle={{
              backgroundColor: "#ffffff",
              border: "1px solid #E5E7EB",
              borderRadius: "8px",
              color: "#111827",
              fontSize: 12,
            }}
          />
          <Legend wrapperStyle={{ fontSize: 11 }} iconSize={10} />
          {segments.map((seg, i) => (
            <Area
              key={seg}
              type="monotone"
              dataKey={seg}
              stackId="1"
              stroke={COLORS[i % COLORS.length]}
              fill={COLORS[i % COLORS.length]}
              fillOpacity={0.75}
              strokeWidth={1}
              isAnimationActive={false}
            />
          ))}
          {!isShare && (
            <Line
              type="monotone"
              dataKey={columns[0]}
              stroke="#111827"
              strokeWidth={1.5}
              strokeDasharray="4 3"
              dot={false}
              isAnimationActive={false}
            />
          )}
        </ComposedChart>
      </ResponsiveContainer>
    </div>
  );
}
