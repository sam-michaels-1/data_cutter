import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  Cell,
  LabelList,
} from "recharts";
import type { WaterfallData } from "../../types/dashboard";
import { formatCurrency } from "../../utils/format";

interface Props {
  waterfall: WaterfallData;
  scaleFactor: number;
}

interface WaterfallEntry {
  name: string;
  range: [number, number];
  type: "total" | "positive" | "negative";
  value: number;
}

export default function WaterfallChart({ waterfall, scaleFactor }: Props) {
  const { bop, new_logo, upsell, downsell, churn, eop } = waterfall;

  // Order: BoP → Churn → Downsell → Upsell → New → EoP
  let running = bop;

  const churnTop = running;
  running += churn; // churn is negative
  const churnBottom = running;

  const downsellTop = running;
  running += downsell; // downsell is negative
  const downsellBottom = running;

  const upsellBottom = running;
  running += upsell;
  const upsellTop = running;

  const newBottom = running;
  running += new_logo;
  const newTop = running;

  const data: WaterfallEntry[] = [
    { name: "BoP", range: [0, bop], type: "total", value: bop },
    { name: "Churn", range: [churnBottom, churnTop], type: "negative", value: churn },
    { name: "Downsell", range: [downsellBottom, downsellTop], type: "negative", value: downsell },
    { name: "Upsell", range: [upsellBottom, upsellTop], type: "positive", value: upsell },
    { name: "New", range: [newBottom, newTop], type: "positive", value: new_logo },
    { name: "EoP", range: [0, eop], type: "total", value: eop },
  ];

  const colors: Record<string, string> = {
    total: "#14B8A6",
    positive: "#34D399",
    negative: "#F87171",
  };

  const renderLabel = (props: any) => {
    const { x, y, width, height, index } = props;
    if (index == null || !data[index]) return null;
    const entry = data[index];
    const isNegative = entry.value < 0;
    if (entry.value === 0) return null;

    const formatted = formatCurrency(entry.value, scaleFactor);

    const labelY = isNegative
      ? (y as number) + (height as number) + 12
      : (y as number) - 5;

    return (
      <text
        x={(x as number) + (width as number) / 2}
        y={labelY}
        textAnchor="middle"
        fill={isNegative ? "#EF4444" : "#374151"}
        fontSize={10}
        fontWeight={500}
      >
        {formatted}
      </text>
    );
  };

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4">
      <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-3">
        Latest Period Waterfall
        <span className="ml-2 text-xs font-normal text-gray-500">
          ({waterfall.period_label})
        </span>
      </h3>
      <ResponsiveContainer width="100%" height={300}>
        <BarChart data={data} margin={{ top: 20, right: 20, bottom: 5, left: 10 }}>
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
            formatter={(value: any) => {
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
            <LabelList content={renderLabel} />
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
