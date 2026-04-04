import type { CohortData, CohortEntry, CohortMetric } from "../../types/dashboard";
import { formatCurrency } from "../../utils/format";

interface Props {
  cohort: CohortData;
  metric: CohortMetric;
  scaleFactor: number;
  granularity: string;
  metricLabel?: string;
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function getMetricValues(c: CohortEntry, metric: CohortMetric): (number | null)[] {
  switch (metric) {
    case "ndr": return c.ndr;
    case "logo_retention": return c.logo_retention;
    case "arr": return c.arr;
    case "customers": return c.customers;
  }
}

/** Shift values left so first non-null becomes index 0 (Y0). */
function leftAlign(values: (number | null)[]): (number | null)[] {
  const first = values.findIndex((v) => v != null);
  return first >= 0 ? values.slice(first) : [];
}

function median(arr: number[]): number {
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

function getCellColor(value: number | null, metric: CohortMetric): string {
  if (value == null) return "";
  if (metric === "ndr" || metric === "logo_retention") {
    if (value >= 1.2) return "bg-emerald-200 text-emerald-800";
    if (value >= 1.0) return "bg-emerald-100 text-emerald-700";
    if (value >= 0.9) return "bg-yellow-100 text-yellow-700";
    if (value >= 0.8) return "bg-orange-100 text-orange-700";
    return "bg-red-100 text-red-700";
  }
  return "bg-teal-50 text-teal-700";
}

function formatValue(value: number | null, metric: CohortMetric, scaleFactor: number): string {
  if (value == null) return "";
  if (metric === "ndr" || metric === "logo_retention") {
    const isNeg = value < 0;
    const formatted = `${(Math.abs(value) * 100).toFixed(0)}%`;
    return isNeg ? `(${formatted})` : formatted;
  }
  if (metric === "arr") {
    return value === 0 ? "$0" : formatCurrency(value, scaleFactor);
  }
  // customers
  return value === 0 ? "0" : value.toLocaleString();
}

// ── Component ─────────────────────────────────────────────────────────────────

const PERIOD_PREFIX: Record<string, string> = {
  annual: "Y",
  quarterly: "Q",
  monthly: "M",
};

export default function CohortHeatmap({ cohort, metric, scaleFactor, granularity, metricLabel = "ARR" }: Props) {
  const { cohorts } = cohort;

  if (!cohorts?.length) {
    return <div className="text-center text-gray-500 py-8">No cohort data available.</div>;
  }

  // Build left-aligned rows
  const aligned = cohorts.map((c) => ({
    c,
    shifted: leftAlign(getMetricValues(c, metric)),
  }));

  // Max number of relative periods across all cohorts
  const maxCols = Math.max(...aligned.map((a) => a.shifted.length));
  const prefix = PERIOD_PREFIX[granularity] ?? "Y";
  const headers = Array.from({ length: maxCols }, (_, i) => `${prefix}${i}`);

  // Summary rows: per column across cohorts
  const avgRow: (number | null)[] = [];
  const medianRow: (number | null)[] = [];
  const dwAvgRow: (number | null)[] = [];

  for (let col = 0; col < maxCols; col++) {
    const vals: number[] = [];
    const weights: number[] = [];

    for (const { c, shifted } of aligned) {
      const v = shifted[col];
      if (v != null) {
        vals.push(v);
        weights.push(c.starting_arr > 0 ? c.starting_arr : 0);
      }
    }

    if (vals.length === 0) {
      avgRow.push(null);
      medianRow.push(null);
      dwAvgRow.push(null);
    } else {
      avgRow.push(vals.reduce((a, b) => a + b, 0) / vals.length);
      medianRow.push(median(vals));
      const totalWeight = weights.reduce((a, b) => a + b, 0);
      if (totalWeight > 0) {
        const wSum = vals.reduce((s, v, i) => s + v * weights[i], 0);
        dwAvgRow.push(wSum / totalWeight);
      } else {
        dwAvgRow.push(vals.reduce((a, b) => a + b, 0) / vals.length);
      }
    }
  }

  // "#" column header and value depend on metric
  const isPctMetric = metric === "ndr" || metric === "logo_retention";
  const sizeHeader = isPctMetric || metric === "arr" ? `Starting $${metricLabel}` : "# Customers";

  return (
    <div className="overflow-x-auto">
      <table className="text-xs border-collapse">
        <thead>
          <tr>
            <th className="px-3 py-2 text-left text-gray-500 font-semibold uppercase tracking-wide sticky left-0 bg-white z-10 whitespace-nowrap">
              Cohort
            </th>
            <th className="px-2 py-2 text-right text-gray-500 font-semibold whitespace-nowrap">
              {sizeHeader}
            </th>
            {headers.map((h) => (
              <th
                key={h}
                className="px-3 py-2 text-center text-gray-500 font-medium whitespace-nowrap"
              >
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {aligned.map(({ c, shifted }) => (
            <tr key={c.label} className="border-t border-gray-200">
              <td className="px-3 py-1.5 font-medium text-gray-800 sticky left-0 bg-white z-10 whitespace-nowrap">
                {c.label}
              </td>
              <td className="px-2 py-1.5 text-right text-gray-500 font-mono whitespace-nowrap">
                {isPctMetric || metric === "arr"
                  ? formatCurrency(c.starting_arr, scaleFactor)
                  : c.count.toLocaleString()}
              </td>
              {headers.map((_, i) => {
                const val = shifted[i] ?? null;
                return (
                  <td
                    key={i}
                    className={`px-3 py-1.5 text-center font-mono rounded-sm ${getCellColor(val, metric)}`}
                  >
                    {formatValue(val, metric, scaleFactor)}
                  </td>
                );
              })}
            </tr>
          ))}

          {/* Summary rows — only for pct metrics */}
          {(metric === "ndr" || metric === "logo_retention") && (
            <>
              <tr className="border-t-2 border-gray-300">
                <td colSpan={2} className="px-3 py-1.5 font-semibold text-gray-600 sticky left-0 bg-white z-10 italic">
                  Average
                </td>
                {avgRow.map((v, i) => (
                  <td key={i} className={`px-3 py-1.5 text-center font-mono italic ${getCellColor(v, metric)}`}>
                    {formatValue(v, metric, scaleFactor)}
                  </td>
                ))}
              </tr>
              <tr className="border-t border-gray-200">
                <td colSpan={2} className="px-3 py-1.5 font-semibold text-gray-600 sticky left-0 bg-white z-10 italic">
                  Median
                </td>
                {medianRow.map((v, i) => (
                  <td key={i} className={`px-3 py-1.5 text-center font-mono italic ${getCellColor(v, metric)}`}>
                    {formatValue(v, metric, scaleFactor)}
                  </td>
                ))}
              </tr>
              <tr className="border-t border-gray-200">
                <td colSpan={2} className="px-3 py-1.5 font-bold text-gray-800 sticky left-0 bg-white z-10 italic">
                  Dollar-Weighted Avg
                </td>
                {dwAvgRow.map((v, i) => (
                  <td key={i} className={`px-3 py-1.5 text-center font-mono font-bold italic ${getCellColor(v, metric)}`}>
                    {formatValue(v, metric, scaleFactor)}
                  </td>
                ))}
              </tr>
            </>
          )}
        </tbody>
      </table>
    </div>
  );
}
