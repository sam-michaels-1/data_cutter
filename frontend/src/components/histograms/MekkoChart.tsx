import { useState, useMemo } from "react";
import type { MekkoData } from "../../engine/histograms";
import { formatCurrency } from "../../utils/format";

interface Props {
  data: MekkoData;
  title: string;
  scaleFactor: number;
  valueType: "arr" | "count";
  metricLabel: string;
}

const COLORS = [
  "#14B8A6", "#F59E0B", "#6366F1", "#EC4899", "#10B981",
  "#8B5CF6", "#F97316", "#06B6D4", "#EF4444", "#84CC16",
  "#A855F7", "#0EA5E9", "#F43F5E", "#22D3EE", "#D946EF",
];

export default function MekkoChart({ data, title, scaleFactor, valueType, metricLabel }: Props) {
  const [hovered, setHovered] = useState<{ col: number; row: number } | null>(null);

  const yColorMap = useMemo(() => {
    const map = new Map<string, string>();
    data.yLabels.forEach((label, i) => {
      map.set(label, COLORS[i % COLORS.length]);
    });
    return map;
  }, [data.yLabels]);

  const formatValue = (v: number) => valueType === "arr" ? formatCurrency(v, scaleFactor) : v.toLocaleString();

  // Compute Y-axis totals for summary table
  const yTotals = useMemo(() => {
    const totals = new Map<string, number>();
    let grandTotal = 0;
    for (const col of data.columns) {
      for (const stack of col.stacks) {
        totals.set(stack.yLabel, (totals.get(stack.yLabel) || 0) + stack.value);
        grandTotal += stack.value;
      }
    }
    return {
      grandTotal,
      rows: data.yLabels.map(label => ({
        label,
        total: totals.get(label) || 0,
        pct: grandTotal > 0 ? (totals.get(label) || 0) / grandTotal : 0,
      })),
    };
  }, [data]);

  // Compute tooltip position outside the overflow-hidden chart body
  const tooltipInfo = useMemo(() => {
    if (!hovered) return null;
    const col = data.columns[hovered.col];
    if (!col) return null;
    const stack = col.stacks[hovered.row];
    if (!stack || stack.value <= 0) return null;

    // Horizontal: center of the hovered column
    const colWidths = data.columns.map(c => Math.max(c.xPct * 100, 2));
    const totalW = colWidths.reduce((s, w) => s + w, 0);
    let cumulativeW = 0;
    for (let i = 0; i < hovered.col; i++) cumulativeW += colWidths[i];
    const leftPct = ((cumulativeW + colWidths[hovered.col] / 2) / totalW) * 100;

    // Vertical: top edge of hovered segment (measured from bottom)
    // flex-col-reverse means stacks[0] is at bottom, stacks[n-1] at top
    let bottomPct = 0;
    for (let i = 0; i < hovered.row; i++) {
      if (col.stacks[i].value > 0) bottomPct += col.stacks[i].pct * 100;
    }
    bottomPct += stack.pct * 100; // top edge of this segment

    // Determine tooltip alignment to prevent clipping at edges
    let translateX = '-50%'; // default: centered
    if (leftPct < 15) translateX = '0%';        // near left edge: align left
    else if (leftPct > 85) translateX = '-100%'; // near right edge: align right

    return { col, stack, leftPct, bottomPct, translateX };
  }, [hovered, data]);

  if (data.columns.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-xl p-3">
        <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-2">{title}</h3>
        <p className="text-sm text-gray-400 py-8 text-center">No data available</p>
      </div>
    );
  }

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-3">
      <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-2">{title}</h3>

      {/* Legend */}
      {data.yLabels.length > 1 && (
        <div className="flex flex-wrap gap-3 mb-3">
          {data.yLabels.map(label => (
            <div key={label} className="flex items-center gap-1.5">
              <div className="w-3 h-3 rounded-sm" style={{ backgroundColor: yColorMap.get(label) }} />
              <span className="text-xs text-gray-600">{label}</span>
            </div>
          ))}
        </div>
      )}

      {/* Chart body + tooltip wrapper */}
      <div className="relative">
        <div className="relative">
          <div className="flex h-[220px] border border-gray-200 rounded-md overflow-hidden">
            {data.columns.map((col, ci) => {
              const widthPct = Math.max(col.xPct * 100, 2);
              return (
                <div
                  key={ci}
                  className="flex flex-col-reverse relative border-r border-gray-100 last:border-r-0"
                  style={{ width: `${widthPct}%` }}
                >
                  {col.stacks.map((stack, si) => {
                    if (stack.value <= 0) return null;
                    const heightPct = stack.pct * 100;
                    const isHovered = hovered?.col === ci && hovered?.row === si;
                    return (
                      <div
                        key={si}
                        className="relative transition-opacity overflow-hidden"
                        style={{
                          height: `${heightPct}%`,
                          backgroundColor: yColorMap.get(stack.yLabel) || COLORS[0],
                          opacity: hovered && !isHovered ? 0.5 : 1,
                        }}
                        onMouseEnter={() => setHovered({ col: ci, row: si })}
                        onMouseLeave={() => setHovered(null)}
                      >
                        {/* Label inside segment if large enough */}
                        {heightPct > 10 && widthPct > 6 && (
                          <span className="absolute inset-0 flex items-center justify-center text-white text-[10px] font-medium drop-shadow-sm">
                            {(stack.pct * 100).toFixed(0)}%
                          </span>
                        )}
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>

          {/* Tooltip - rendered outside overflow-hidden chart body */}
          {tooltipInfo && (
            <div
              className="absolute z-20 bg-white border border-gray-200 rounded-lg shadow-lg px-3 py-2 text-xs pointer-events-none whitespace-nowrap"
              style={{
                left: `${tooltipInfo.leftPct}%`,
                bottom: `${tooltipInfo.bottomPct}%`,
                transform: `translateX(${tooltipInfo.translateX}) translateY(-4px)`,
              }}
            >
              <div className="font-medium text-gray-800">{tooltipInfo.stack.yLabel}</div>
              <div className="text-gray-600">
                {valueType === "arr" ? metricLabel : "Customers"}: {formatValue(tooltipInfo.stack.value)}
              </div>
              <div className="text-gray-500">
                {(tooltipInfo.stack.pct * 100).toFixed(1)}% of {tooltipInfo.col.xLabel}
              </div>
            </div>
          )}
        </div>

        {/* X-axis labels */}
        <div className="flex mt-1 overflow-x-clip">
          {data.columns.map((col, ci) => {
            const widthPct = Math.max(col.xPct * 100, 2);
            return (
              <div key={ci} className="text-center" style={{ width: `${widthPct}%` }}>
                <div className="text-[10px] text-gray-600 font-medium whitespace-nowrap px-0.5">{col.xLabel}</div>
                <div className="text-[9px] text-gray-400 whitespace-nowrap px-0.5">
                  {formatValue(col.xTotal)} / {(col.xPct * 100).toFixed(1)}%
                </div>
              </div>
            );
          })}
        </div>

        {/* Y-Axis Summary Table */}
        {data.yLabels.length > 1 && (
          <div className="mt-3">
            <table className="w-full text-xs">
              <thead>
                <tr className="border-b border-gray-200">
                  <th className="text-left py-1.5 px-2 text-gray-500 font-medium">Category</th>
                  <th className="text-right py-1.5 px-2 text-gray-500 font-medium">
                    {valueType === "arr" ? metricLabel : "Customers"}
                  </th>
                  <th className="text-right py-1.5 px-2 text-gray-500 font-medium">% of Total</th>
                </tr>
              </thead>
              <tbody>
                {yTotals.rows.map(({ label, total, pct }) => (
                  <tr key={label} className="border-b border-gray-50">
                    <td className="py-1.5 px-2 text-gray-700 font-medium">
                      <span className="inline-flex items-center gap-1.5">
                        <span
                          className="w-2.5 h-2.5 rounded-sm inline-block flex-shrink-0"
                          style={{ backgroundColor: yColorMap.get(label) }}
                        />
                        {label}
                      </span>
                    </td>
                    <td className="text-right py-1.5 px-2 text-gray-600 font-mono">
                      {formatValue(total)}
                    </td>
                    <td className="text-right py-1.5 px-2 text-gray-500">
                      {(pct * 100).toFixed(1)}%
                    </td>
                  </tr>
                ))}
                <tr className="border-t border-gray-200">
                  <td className="py-1.5 px-2 text-gray-700 font-semibold">Total</td>
                  <td className="text-right py-1.5 px-2 text-gray-700 font-mono font-semibold">
                    {formatValue(yTotals.grandTotal)}
                  </td>
                  <td className="text-right py-1.5 px-2 text-gray-500 font-semibold">100.0%</td>
                </tr>
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
