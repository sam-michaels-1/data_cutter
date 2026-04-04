import type { GridData } from "../../engine/histograms";

interface Props {
  data: GridData;
  title: string;
  subtitle?: string;
  formatMetric: (v: number) => string;
  colorScale: (v: number) => { bg: string; text: string };
}

export default function TwoByTwoGrid({ data, title, subtitle, formatMetric, colorScale }: Props) {
  if (data.xLabels.length === 0 || data.yLabels.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-xl p-4">
        <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-3">{title}</h3>
        <p className="text-sm text-gray-400 py-8 text-center">Not enough data</p>
      </div>
    );
  }

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4">
      <div className="mb-3">
        <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide">{title}</h3>
        {subtitle && <p className="text-[10px] text-gray-400 mt-0.5">{subtitle}</p>}
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-xs">
          <thead>
            <tr>
              <th className="text-left py-2 pr-3 text-gray-500 font-medium" />
              {data.xLabels.map(x => (
                <th key={x} className="text-center py-2 px-2 text-gray-500 font-medium whitespace-nowrap">{x}</th>
              ))}
              <th className="text-center py-2 px-2 text-gray-500 font-medium bg-gray-50">Total</th>
            </tr>
          </thead>
          <tbody>
            {data.yLabels.map((yLabel, yi) => (
              <tr key={yLabel} className="border-t border-gray-100">
                <td className="py-2 pr-3 text-gray-600 font-medium whitespace-nowrap">{yLabel}</td>
                {data.xLabels.map((_, xi) => {
                  const cell = data.grid[yi]?.[xi];
                  if (!cell || cell.value == null) {
                    return <td key={xi} className="text-center py-2 px-2 text-gray-300">-</td>;
                  }
                  const colors = colorScale(cell.value);
                  return (
                    <td
                      key={xi}
                      className="text-center py-2 px-2 font-mono font-medium"
                      style={{ backgroundColor: colors.bg, color: colors.text }}
                    >
                      <div>{formatMetric(cell.value)}</div>
                      <div className="text-[9px] text-gray-400 font-normal">n={cell.count}</div>
                    </td>
                  );
                })}
                <td className="text-center py-2 px-2 font-mono font-medium bg-gray-50">
                  {data.yTotals[yi] != null ? formatMetric(data.yTotals[yi]!) : '-'}
                </td>
              </tr>
            ))}
            {/* Column totals row */}
            <tr className="border-t-2 border-gray-200 bg-gray-50">
              <td className="py-2 pr-3 text-gray-600 font-medium">Total</td>
              {data.xLabels.map((_, xi) => (
                <td key={xi} className="text-center py-2 px-2 font-mono font-medium text-gray-700">
                  {data.xTotals[xi] != null ? formatMetric(data.xTotals[xi]!) : '-'}
                </td>
              ))}
              <td className="text-center py-2 px-2" />
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}
