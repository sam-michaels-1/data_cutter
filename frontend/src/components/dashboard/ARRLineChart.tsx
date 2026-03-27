import {
  ResponsiveContainer,
  LineChart,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
} from "recharts";

interface Props {
  periods: string[];
  arrOverTime: number[];
}

function formatYAxis(value: number): string {
  if (value >= 1_000_000) return `$${(value / 1_000_000).toFixed(1)}M`;
  if (value >= 1_000) return `$${(value / 1_000).toFixed(0)}K`;
  return `$${value}`;
}

export default function ARRLineChart({ periods, arrOverTime }: Props) {
  const data = periods.map((label, i) => ({
    period: label,
    arr: arrOverTime[i],
  }));

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4">
      <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-3">
        ARR Over Time
      </h3>
      <ResponsiveContainer width="100%" height={260}>
        <LineChart data={data} margin={{ top: 5, right: 20, bottom: 5, left: 10 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#E5E7EB" opacity={0.8} />
          <XAxis
            dataKey="period"
            tick={{ fontSize: 11, fill: "#6B7280" }}
            axisLine={{ stroke: "#D1D5DB" }}
            tickLine={false}
          />
          <YAxis
            tickFormatter={formatYAxis}
            tick={{ fontSize: 11, fill: "#6B7280" }}
            axisLine={false}
            tickLine={false}
          />
          <Tooltip
            formatter={(value: number) => [formatYAxis(value), "ARR"]}
            contentStyle={{
              backgroundColor: "#ffffff",
              border: "1px solid #E5E7EB",
              borderRadius: "8px",
              color: "#111827",
              fontSize: 12,
            }}
          />
          <Line
            type="monotone"
            dataKey="arr"
            stroke="#14B8A6"
            strokeWidth={2.5}
            dot={{ fill: "#14B8A6", r: 3 }}
            activeDot={{ r: 5 }}
          />
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}
