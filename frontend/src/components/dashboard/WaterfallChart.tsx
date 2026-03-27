import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  Cell,
} from "recharts";
import type { WaterfallData } from "../../types/dashboard";
import { formatCurrency } from "../../utils/format";

interface Props {
  waterfall: WaterfallData;
  scaleFactor: number;
}

export default function WaterfallChart({ waterfall, scaleFactor }: Props) {
  const { bop, new_logo, upsell, downsell, churn, eop } = waterfall;

  let running = bop;

  const newBottom = running;
  running += new_logo;
  const newTop = running;

  const upsellBottom = running;
  running += upsell;
  const upsellTop = running;

  const downsellTop = running;
  running += downsell;
  const downsellBottom = running;

  const churnTop = running;
  running += churn;
  const churnBottom = running;

  const data = [
    { name: "BoP", range: [0, bop] as [number, number], type: "total" },
    { name: "New", range: [newBottom, newTop] as [number, number], type: "positive" },
    { name: "Upsell", range: [upsellBottom, upsellTop] as [number, number], type: "positive" },
    { name: "Downsell", range: [downsellBottom, downsellTop] as [number, number], type: "negative" },
    { name: "Churn", range: [churnBottom, churnTop] as [number, number], type: "negative" },
    { name: "EoP", range: [0, eop] as [number, number], type: "total" },
  ];

  const colors: Record<string, string> = {
    total: "#14B8A6",
    positive: "#34D399",
    negative: "#F87171",
  };

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4">
      <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-3">
        Latest Period Waterfall
        <span className="ml-2 text-xs font-normal text-gray-500">
          ({waterfall.period_label})
        </span>
      </h3>
      <ResponsiveContainer width="100%" height={280}>
        <BarChart data={data} margin={{ top: 5, right: 20, bottom: 5, left: 10 }}>
          <XAxis
            dataKey="name"
            tick={{ fontSize: 11, fill: "#6B7280" }}
            axisLine={{ stroke: "#D1D5DB" }}
            tickLine={false}
          />
          <YAxis
            tickFormatter={(v) => formatCurrency(v, scaleFactor)}
            tick={{ fontSize: 11, fill: "#6B7280" }}
            axisLine={false}
            tickLine={false}
          />
          <Tooltip
            formatter={(value: [number, number]) => {
              const diff = Math.round((value[1] - value[0]) * 10) / 10;
              return [formatCurrency(Math.abs(diff), scaleFactor), "Amount"];
            }}
            contentStyle={{
              backgroundColor: "#ffffff",
              border: "1px solid #E5E7EB",
              borderRadius: "8px",
              color: "#111827",
              fontSize: 12,
            }}
          />
          <Bar dataKey="range" radius={[3, 3, 0, 0]} isAnimationActive={false}>
            {data.map((entry, idx) => (
              <Cell key={idx} fill={colors[entry.type]} />
            ))}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
