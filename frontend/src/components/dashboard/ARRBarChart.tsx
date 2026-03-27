import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  LabelList,
} from "recharts";
import { formatCurrency, formatPct } from "../../utils/format";

interface Props {
  periods: string[];
  arrOverTime: number[];
  arrGrowthPcts: (number | null)[];
  scaleFactor: number;
}

// Custom XAxis tick that renders period label + growth % below it
function CustomTick({ x, y, payload, growthPcts, periods }: {
  x?: number;
  y?: number;
  payload?: { value: string };
  growthPcts: (number | null)[];
  periods: string[];
}) {
  const idx = periods.indexOf(payload?.value ?? "");
  const growth = idx >= 0 ? growthPcts[idx] : null;

  return (
    <g transform={`translate(${x},${y})`}>
      <text x={0} y={0} dy={12} textAnchor="middle" fill="#6B7280" fontSize={11}>
        {payload?.value}
      </text>
      {growth != null && (
        <text
          x={0}
          y={0}
          dy={24}
          textAnchor="middle"
          fill={growth >= 0 ? "#10B981" : "#EF4444"}
          fontSize={9}
          fontWeight={500}
        >
          {growth >= 0 ? "+" : ""}{(growth * 100).toFixed(1)}%
        </text>
      )}
    </g>
  );
}

export default function ARRBarChart({ periods, arrOverTime, arrGrowthPcts, scaleFactor }: Props) {
  const data = periods.map((label, i) => ({
    period: label,
    arr: arrOverTime[i],
  }));

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4">
      <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-3">
        ARR Over Time
      </h3>
      <ResponsiveContainer width="100%" height={280}>
        <BarChart data={data} margin={{ top: 5, right: 20, bottom: 30, left: 10 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#E5E7EB" opacity={0.8} vertical={false} />
          <XAxis
            dataKey="period"
            tick={(props) => (
              <CustomTick
                {...props}
                growthPcts={arrGrowthPcts}
                periods={periods}
              />
            )}
            axisLine={{ stroke: "#D1D5DB" }}
            tickLine={false}
            interval={0}
          />
          <YAxis
            tickFormatter={(v) => formatCurrency(v, scaleFactor)}
            tick={{ fontSize: 11, fill: "#6B7280" }}
            axisLine={false}
            tickLine={false}
          />
          <Tooltip
            formatter={(value: number) => [formatCurrency(value, scaleFactor), "ARR"]}
            contentStyle={{
              backgroundColor: "#ffffff",
              border: "1px solid #E5E7EB",
              borderRadius: "8px",
              color: "#111827",
              fontSize: 12,
            }}
          />
          <Bar dataKey="arr" fill="#14B8A6" radius={[3, 3, 0, 0]} isAnimationActive={false} />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
