import type { StatsData } from "../../types/dashboard";
import { formatCurrency, formatPct } from "../../utils/format";

interface Props {
  stats: StatsData;
  scaleFactor: number;
  latestPeriodLabel: string;
  latestPeriodDate: string;
  metricLabel?: string;
}

const CARDS: {
  key: keyof StatsData;
  label: string;
  format: "currency" | "pct" | "count";
}[] = [
  { key: "total_arr", label: "Total {metric}", format: "currency" },
  { key: "customer_count", label: "Customers", format: "count" },
  { key: "yoy_growth_pct", label: "YoY Growth", format: "pct" },
  { key: "punitive_retention_pct", label: "Punitive Retention", format: "pct" },
  { key: "lost_only_retention_pct", label: "Lost-Only Retention", format: "pct" },
  { key: "net_retention_pct", label: "Net Retention", format: "pct" },
];

function formatAsOfDate(isoDate: string): string {
  if (!isoDate) return "";
  try {
    const d = new Date(isoDate + "T00:00:00"); // force local-time parse
    return d.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
  } catch {
    return isoDate;
  }
}

export default function StatsCards({ stats, scaleFactor, latestPeriodLabel, latestPeriodDate, metricLabel = "ARR" }: Props) {
  return (
    <div className="space-y-2">
      {latestPeriodLabel && (
        <p className="text-xs text-gray-500">
          As of <span className="font-medium">{latestPeriodLabel}</span>
          {latestPeriodDate && (
            <span> ({formatAsOfDate(latestPeriodDate)})</span>
          )}
        </p>
      )}
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
        {CARDS.map(({ key, label: rawLabel, format }) => {
          const label = rawLabel.replace("{metric}", metricLabel);
          const raw = stats[key];
          let display: string;
          let color = "text-gray-900";

          if (format === "currency") {
            display = formatCurrency(raw as number, scaleFactor);
          } else if (format === "pct") {
            display = formatPct(raw as number | null);
            if (raw != null) {
              color =
                (raw as number) >= 1
                  ? "text-emerald-600"
                  : (raw as number) >= 0
                  ? "text-amber-500"
                  : "text-red-500";
            }
          } else {
            display = (raw as number).toLocaleString();
          }

          return (
            <div
              key={key}
              className="bg-white border border-gray-200 rounded-xl px-3 py-2"
            >
              <p className="text-xs font-medium text-gray-500 uppercase tracking-wide">
                {label}
              </p>
              <p className={`text-xl font-bold mt-1 ${color}`}>{display}</p>
            </div>
          );
        })}
      </div>
    </div>
  );
}
