import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  Cell,
} from "recharts";
import type { HistogramBucket } from "../../engine/histograms";

interface Props {
  data: HistogramBucket[];
}

export default function GrowthHistogram({ data }: Props) {
  if (data.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-xl p-4">
        <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-3">
          LTM Growth Rate Distribution
        </h3>
        <p className="text-sm text-gray-400 py-8 text-center">
          Not enough data for growth calculation (need at least 2 comparable periods)
        </p>
      </div>
    );
  }

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4">
      <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-3">
        LTM Growth Rate Distribution
      </h3>
      <ResponsiveContainer width="100%" height={300}>
        <BarChart data={data} margin={{ top: 5, right: 20, bottom: 5, left: 10 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6" />
          <XAxis
            dataKey="label"
            tick={{ fontSize: 10, fill: "#6B7280" }}
            axisLine={{ stroke: "#D1D5DB" }}
            tickLine={false}
            interval={0}
            angle={-20}
            textAnchor="end"
            height={50}
          />
          <YAxis
            tick={{ fontSize: 11, fill: "#6B7280" }}
            axisLine={false}
            tickLine={false}
            allowDecimals={false}
            label={{ value: "# Customers", angle: -90, position: "insideLeft", style: { fontSize: 11, fill: "#9CA3AF" } }}
          />
          <Tooltip
            formatter={(value: unknown) => [String(value), "Customers"]}
            contentStyle={{
              backgroundColor: "#fff",
              border: "1px solid #E5E7EB",
              borderRadius: "8px",
              fontSize: 12,
            }}
          />
          <Bar dataKey="count" radius={[3, 3, 0, 0]}>
            {data.map((entry, idx) => {
              // Color by whether bucket represents positive or negative growth
              const midpoint = (entry.min === -Infinity ? entry.max : entry.max === Infinity ? entry.min : (entry.min + entry.max) / 2);
              const color = midpoint >= 0 ? "#34D399" : "#F87171";
              return <Cell key={idx} fill={color} />;
            })}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
