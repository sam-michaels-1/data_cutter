import { useState } from "react";
import { useSession } from "../components/SessionProvider";
import { useDashboard } from "../hooks/useDashboard";
import CohortHeatmap from "../components/dashboard/CohortHeatmap";
import type { CohortMetric } from "../types/dashboard";

const METRICS: { key: CohortMetric; label: string }[] = [
  { key: "arr", label: "ARR" },
  { key: "ndr", label: "Dollar Retention" },
  { key: "customers", label: "Customers" },
  { key: "logo_retention", label: "Logo Retention" },
];

export default function CohortPage() {
  const { sessionId } = useSession();
  const { data, loading, error, refetch } = useDashboard(sessionId);
  const [metric, setMetric] = useState<CohortMetric>("arr");

  if (!sessionId) {
    return (
      <div className="flex items-center justify-center h-full min-h-[60vh]">
        <div className="text-center text-gray-500">
          <p className="text-lg font-medium">No data imported yet</p>
          <p className="text-sm mt-1">Go to Import to upload your data file.</p>
        </div>
      </div>
    );
  }

  if (loading && !data) {
    return (
      <div className="flex items-center justify-center h-full min-h-[60vh]">
        <div className="text-center text-gray-500">
          <div className="animate-spin h-8 w-8 border-2 border-teal-500 border-t-transparent rounded-full mx-auto mb-3" />
          <p className="text-sm">Loading cohort data...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex items-center justify-center h-full min-h-[60vh]">
        <div className="text-center text-red-500">
          <p>{error}</p>
          <button
            onClick={() => refetch()}
            className="mt-3 px-4 py-2 bg-teal-600 text-white rounded-lg text-sm"
          >
            Retry
          </button>
        </div>
      </div>
    );
  }

  if (!data) return null;

  const { cohort, granularity, available_granularities, scale_factor } = data;

  const title =
    metric === "arr"
      ? "ARR by Cohort"
      : metric === "ndr"
      ? "Net Dollar Retention (%)"
      : metric === "logo_retention"
      ? "Logo Retention (%)"
      : "Customer Count by Cohort";

  return (
    <div className="p-6 space-y-4 max-w-[1400px]">
      <div>
        <h1 className="text-xl font-bold text-gray-900">Cohort Analysis</h1>
        <p className="text-sm text-gray-500">
          Track retention and revenue by customer cohort
        </p>
      </div>

      <div className="flex items-center justify-between flex-wrap gap-3">
        {/* Metric toggle */}
        <div className="flex gap-1 bg-gray-100 rounded-lg p-0.5">
          {METRICS.map(({ key, label }) => (
            <button
              key={key}
              onClick={() => setMetric(key)}
              className={`px-3 py-1 rounded-md text-xs font-medium transition ${
                key === metric
                  ? "bg-teal-600 text-white shadow"
                  : "text-gray-500 hover:text-gray-700"
              }`}
            >
              {label}
            </button>
          ))}
        </div>

        {/* Granularity toggle */}
        {available_granularities.length > 1 && (
          <div className="flex gap-1 bg-gray-100 rounded-lg p-0.5">
            {available_granularities.map((g) => (
              <button
                key={g}
                onClick={() => refetch(g)}
                className={`px-3 py-1 rounded-md text-xs font-medium transition ${
                  g === granularity
                    ? "bg-teal-600 text-white shadow"
                    : "text-gray-500 hover:text-gray-700"
                }`}
              >
                {g.charAt(0).toUpperCase() + g.slice(1)}
              </button>
            ))}
          </div>
        )}
      </div>

      <div className="bg-white border border-gray-200 rounded-xl p-4">
        <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-3">
          {title}
        </h3>
        <CohortHeatmap cohort={cohort} metric={metric} scaleFactor={scale_factor} />
      </div>
    </div>
  );
}
