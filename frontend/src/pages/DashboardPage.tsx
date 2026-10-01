import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { useSession } from "../components/SessionProvider";
import { useDashboard } from "../hooks/useDashboard";
import { useSummaryData } from "../hooks/useSummaryData";
import StatsCards from "../components/dashboard/StatsCards";
import ARRBarChart from "../components/dashboard/ARRBarChart";
import SegmentBarChart from "../components/dashboard/SegmentBarChart";
import WaterfallChart from "../components/dashboard/WaterfallChart";
import TopCustomersTable from "../components/dashboard/TopCustomersTable";
import AttributeFilterBar from "../components/AttributeFilterBar";
import type { Filters } from "../types/dashboard";

export default function DashboardPage() {
  const { sessionId } = useSession();
  const navigate = useNavigate();
  const { data, loading, error, refetch } = useDashboard(sessionId);
  const { data: summaryData, refetch: refetchSummary } = useSummaryData(sessionId);
  const [filters, setFilters] = useState<Filters>({});
  const [segmentBy, setSegmentBy] = useState("");

  const handleGranularityChange = (g: string) => {
    refetch(g, { filters });
    refetchSummary(g, { filters, identifier: segmentBy || undefined });
  };

  const handleFilterChange = (newFilters: Filters) => {
    setFilters(newFilters);
    refetch(data?.granularity, { filters: newFilters });
    refetchSummary(data?.granularity, { filters: newFilters, identifier: segmentBy || undefined });
  };

  const handleSegmentByChange = (v: string) => {
    setSegmentBy(v);
    refetchSummary(data?.granularity, { filters, identifier: v });
  };

  // Keep the segment chart on the dashboard's granularity
  useEffect(() => {
    if (data?.granularity && summaryData && summaryData.granularity !== data.granularity) {
      refetchSummary(data.granularity, { filters, identifier: segmentBy || undefined });
    }
  }, [data?.granularity, summaryData?.granularity, summaryData, refetchSummary, filters, segmentBy]);

  if (!sessionId) {
    return (
      <div className="flex items-center justify-center h-full min-h-[60vh]">
        <div className="text-center text-gray-500">
          <p className="text-lg font-medium">No data imported yet</p>
          <p className="text-sm mt-1">Please return to the Import tab to load your data.</p>
          <button
            onClick={() => navigate("/import")}
            className="mt-3 px-4 py-2 bg-teal-600 text-white rounded-lg text-sm hover:bg-teal-700 transition"
          >
            Go to Import
          </button>
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
        <div className="text-center text-gray-500">
          <p className="text-lg font-medium">No data loaded</p>
          <p className="text-sm mt-1">Please return to the Import tab to load your data.</p>
          <button
            onClick={() => navigate("/import")}
            className="mt-3 px-4 py-2 bg-teal-600 text-white rounded-lg text-sm hover:bg-teal-700 transition"
          >
            Go to Import
          </button>
        </div>
      </div>
    );
  }

  if (!data) return null;

  const { overview, granularity, available_granularities, scale_factor, attribute_options, data_type } = data;
  const metricLabel = data_type === "revenue" ? "Revenue" : "ARR";

  return (
    <div className="p-3 sm:p-4 space-y-3 max-w-[1600px]">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
        <h1 className="text-xl font-bold text-gray-900">Dashboard</h1>
        {available_granularities?.length > 1 && (
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
      {attribute_options?.length > 0 && (
        <AttributeFilterBar
          attributes={attribute_options}
          filters={filters}
          onChange={handleFilterChange}
          className="sticky top-0 z-30 bg-gray-50 py-2 border-b border-gray-200 !mt-0"
        />
      )}

      {/* Stats cards (includes "as of" label) */}
      <StatsCards
        stats={overview.stats}
        scaleFactor={scale_factor}
        latestPeriodLabel={overview.latest_period_label}
        latestPeriodDate={overview.latest_period_date}
        metricLabel={metricLabel}
        granularity={granularity}
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

      {/* Revenue/ARR over time by segment */}
      {(() => {
        const dollars = summaryData?.sections.find(s => s.key === "dollars");
        if (!summaryData || !dollars) return null;
        const segIdentifier = summaryData.identifiers.includes(segmentBy)
          ? segmentBy
          : summaryData.identifier;
        return (
          <SegmentBarChart
            dollars={dollars}
            pctOfTotal={summaryData.sections.find(s => s.key === "pct_of_total")}
            columns={summaryData.columns}
            scaleFactor={scale_factor}
            metricLabel={metricLabel}
            identifier={segIdentifier}
            identifiers={summaryData.identifiers}
            onIdentifierChange={handleSegmentByChange}
          />
        );
      })()}

      {/* Top customers */}
      {overview.top_customers?.length > 0 && (
        <TopCustomersTable
          customers={overview.top_customers}
          scaleFactor={scale_factor}
          metricLabel={metricLabel}
        />
      )}
    </div>
  );
}
