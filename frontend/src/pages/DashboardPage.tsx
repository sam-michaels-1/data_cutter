import { useState } from "react";
import { useSession } from "../components/SessionProvider";
import { useDashboard } from "../hooks/useDashboard";
import StatsCards from "../components/dashboard/StatsCards";
import ARRBarChart from "../components/dashboard/ARRBarChart";
import WaterfallChart from "../components/dashboard/WaterfallChart";
import TopCustomersTable from "../components/dashboard/TopCustomersTable";
import AttributeFilterBar from "../components/AttributeFilterBar";

export default function DashboardPage() {
  const { sessionId } = useSession();
  const { data, loading, error, refetch } = useDashboard(sessionId);
  const [filters, setFilters] = useState<Record<string, string>>({});

  const handleGranularityChange = (g: string) => {
    refetch(g, { filters });
  };

  const handleFilterChange = (newFilters: Record<string, string>) => {
    setFilters(newFilters);
    refetch(data?.granularity, { filters: newFilters });
  };

  if (!sessionId) {
    return (
      <div className="flex items-center justify-center h-full min-h-[60vh]">
        <div className="text-center text-gray-500 text-gray-500">
          <p className="text-lg font-medium">No data imported yet</p>
          <p className="text-sm mt-1">
            Go to Import to upload your data file.
          </p>
        </div>
      </div>
    );
  }

  if (loading && !data) {
    return (
      <div className="flex items-center justify-center h-full min-h-[60vh]">
        <div className="text-center text-gray-500 text-gray-500">
          <div className="animate-spin h-8 w-8 border-2 border-teal-500 border-t-transparent rounded-full mx-auto mb-3" />
          <p className="text-sm">Computing dashboard...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex items-center justify-center h-full min-h-[60vh]">
        <div className="text-center text-red-500">
          <p className="text-lg font-medium">Error</p>
          <p className="text-sm mt-1">{error}</p>
          <button
            onClick={() => refetch()}
            className="mt-3 px-4 py-2 bg-teal-600 text-white rounded-lg text-sm hover:bg-teal-700 transition"
          >
            Retry
          </button>
        </div>
      </div>
    );
  }

  if (!data) return null;

  const { overview, granularity, available_granularities, scale_factor, attribute_options, data_type } = data;
  const metricLabel = data_type === "revenue" ? "Revenue" : "ARR";

  return (
    <div className="p-6 space-y-4 max-w-[1400px]">
      {/* Header */}
      <div className="flex items-center justify-between">
        <h1 className="text-xl font-bold text-gray-900">Dashboard</h1>
        {available_granularities.length > 1 && (
          <div className="flex gap-1 bg-gray-100 rounded-lg p-0.5">
            {available_granularities.map((g) => (
              <button
                key={g}
                onClick={() => handleGranularityChange(g)}
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

      {/* Attribute filters */}
      {attribute_options.length > 0 && (
        <AttributeFilterBar
          attributes={attribute_options}
          filters={filters}
          onChange={handleFilterChange}
        />
      )}

      {/* Stats cards (includes "as of" label) */}
      <StatsCards
        stats={overview.stats}
        scaleFactor={scale_factor}
        latestPeriodLabel={overview.latest_period_label}
        latestPeriodDate={overview.latest_period_date}
        metricLabel={metricLabel}
      />

      {/* Charts row */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <ARRBarChart
          periods={overview.periods}
          arrOverTime={overview.arr_over_time}
          arrGrowthPcts={overview.arr_growth_pcts}
          scaleFactor={scale_factor}
          metricLabel={metricLabel}
        />
        {overview.waterfall && (
          <WaterfallChart waterfall={overview.waterfall} scaleFactor={scale_factor} />
        )}
      </div>

      {/* Top customers */}
      {overview.top_customers.length > 0 && (
        <TopCustomersTable
          customers={overview.top_customers}
          scaleFactor={scale_factor}
          metricLabel={metricLabel}
        />
      )}
    </div>
  );
}
