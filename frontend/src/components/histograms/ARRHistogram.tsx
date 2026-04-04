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
import type { HistogramBucket } from "../../engine/histograms";

interface Props {
  data: HistogramBucket[];
  metricLabel: string;
}

export default function ARRHistogram({ data, metricLabel }: Props) {
  if (data.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-xl p-3">
        <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-2">
          {metricLabel} Distribution by Customer
        </h3>
        <p className="text-sm text-gray-400 py-8 text-center">No data available</p>
      </div>
    );
  }

  const total = data.reduce((s, d) => s + d.count, 0);

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-3">
      <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-2">
        {metricLabel} Distribution by Customer
      </h3>
      <ResponsiveContainer width="100%" height={220}>
        <BarChart data={data} margin={{ top: 20, right: 20, bottom: 5, left: 10 }}>
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
          <Bar dataKey="count" fill="#14B8A6" radius={[3, 3, 0, 0]}>
            <LabelList
              dataKey="count"
              position="top"
              formatter={(value) =>
                total > 0 ? `${((Number(value) / total) * 100).toFixed(1)}%` : ''
              }
              style={{ fontSize: 10, fill: '#6B7280' }}
            />
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
