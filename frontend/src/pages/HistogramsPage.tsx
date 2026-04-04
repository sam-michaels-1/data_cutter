import { useState, useCallback } from "react";
import { useSession } from "../components/SessionProvider";
import { useHistogramData } from "../hooks/useHistogramData";
import AttributeFilterBar from "../components/AttributeFilterBar";
import ARRHistogram from "../components/histograms/ARRHistogram";
import GrowthHistogram from "../components/histograms/GrowthHistogram";
import MekkoChart from "../components/histograms/MekkoChart";
import IdentifierPieCharts from "../components/histograms/IdentifierPieCharts";
import TwoByTwoGrid from "../components/histograms/TwoByTwoGrid";
import RetentionGrids from "../components/histograms/RetentionGrids";
import type { Filters } from "../types/dashboard";

function AxisSelector({ label, value, options, onChange }: {
  label: string;
  value: string;
  options: string[];
  onChange: (v: string) => void;
}) {
  return (
    <div className="flex items-center gap-1.5">
      <span className="text-xs text-gray-500 font-medium">{label}:</span>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="text-xs border border-gray-200 rounded-md px-2 py-1 bg-white text-gray-700 focus:outline-none focus:ring-1 focus:ring-teal-500"
      >
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );
}

export default function HistogramsPage() {
  const { sessionId } = useSession();
  const { data, loading, error, refetch } = useHistogramData(sessionId);
  const [filters, setFilters] = useState<Filters>({});
  const [mekkoXAxis, setMekkoXAxis] = useState("Cohort");
  const [mekkoYAxis, setMekkoYAxis] = useState("");
  const [gridXAxis, setGridXAxis] = useState("Cohort");
  const [gridYAxis, setGridYAxis] = useState("");

  const currentGran = data?.granularity;

  const doRefetch = useCallback((
    gran?: string,
    opts?: { filters?: Filters; mx?: string; my?: string; gx?: string; gy?: string }
  ) => {
    refetch(gran || currentGran, {
      filters: opts?.filters ?? filters,
      mekkoXAxis: opts?.mx ?? mekkoXAxis,
      mekkoYAxis: opts?.my ?? mekkoYAxis,
      gridXAxis: opts?.gx ?? gridXAxis,
      gridYAxis: opts?.gy ?? gridYAxis,
    });
  }, [refetch, currentGran, filters, mekkoXAxis, mekkoYAxis, gridXAxis, gridYAxis]);

  const handleGranularityChange = (g: string) => doRefetch(g);
  const handleFilterChange = (f: Filters) => { setFilters(f); doRefetch(undefined, { filters: f }); };
  const handleMekkoXChange = (v: string) => { setMekkoXAxis(v); doRefetch(undefined, { mx: v }); };
  const handleMekkoYChange = (v: string) => { setMekkoYAxis(v); doRefetch(undefined, { my: v }); };
  const handleGridXChange = (v: string) => { setGridXAxis(v); doRefetch(undefined, { gx: v }); };
  const handleGridYChange = (v: string) => { setGridYAxis(v); doRefetch(undefined, { gy: v }); };

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
          <p className="text-sm">Computing distributions...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex items-center justify-center h-full min-h-[60vh]">
        <div className="text-center text-red-500">
          <p>{error}</p>
          <button onClick={() => doRefetch()} className="mt-3 px-4 py-2 bg-teal-600 text-white rounded-lg text-sm">
            Retry
          </button>
        </div>
      </div>
    );
  }

  if (!data) return null;

  const { identifiers, available_granularities, granularity, scale_factor, data_type, attribute_options, latestPeriodLabel, priorPeriodLabel } = data;
  const metricLabel = data_type === "revenue" ? "Revenue" : "ARR";
  const periodSubtitle = priorPeriodLabel && latestPeriodLabel
    ? `${priorPeriodLabel} to ${latestPeriodLabel}`
    : latestPeriodLabel ? `As of ${latestPeriodLabel}` : '';

  // Initialize axis defaults from available identifiers
  const effectiveMekkoX = identifiers.includes(mekkoXAxis) ? mekkoXAxis : "Cohort";
  const effectiveMekkoY = mekkoYAxis && identifiers.includes(mekkoYAxis) ? mekkoYAxis : (identifiers.length > 1 ? identifiers.find(i => i !== effectiveMekkoX) || "" : "");
  const effectiveGridX = identifiers.includes(gridXAxis) ? gridXAxis : "Cohort";
  const effectiveGridY = gridYAxis && identifiers.includes(gridYAxis) ? gridYAxis : (identifiers.length > 1 ? identifiers.find(i => i !== effectiveGridX) || "" : "");

  function formatGrowth(v: number): string {
    const formatted = `${(Math.abs(v) * 100).toFixed(1)}%`;
    return v < 0 ? `(${formatted})` : formatted;
  }

  function growthColor(v: number): { bg: string; text: string } {
    if (v >= 0.5)  return { bg: "rgba(5, 150, 105, 0.38)", text: "#047857" };
    if (v >= 0.3)  return { bg: "rgba(5, 150, 105, 0.28)", text: "#059669" };
    if (v >= 0.2)  return { bg: "rgba(16, 185, 129, 0.22)", text: "#059669" };
    if (v >= 0.1)  return { bg: "rgba(16, 185, 129, 0.14)", text: "#10B981" };
    if (v >= 0)    return { bg: "rgba(245, 158, 11, 0.14)", text: "#D97706" };
    if (v >= -0.1) return { bg: "rgba(249, 115, 22, 0.16)", text: "#EA580C" };
    return { bg: "rgba(239, 68, 68, 0.20)", text: "#DC2626" };
  }

  return (
    <div className="p-4 space-y-3 max-w-[1600px]">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-bold text-gray-900">Histograms & Distributions</h1>
          <p className="text-sm text-gray-500">Explore data distributions and segment analysis</p>
        </div>
        {available_granularities.length > 1 && (
          <div className="flex gap-1 bg-gray-100 rounded-lg p-0.5">
            {available_granularities.map(g => (
              <button
                key={g}
                onClick={() => handleGranularityChange(g)}
                className={`px-3 py-1 rounded-md text-xs font-medium transition ${
                  g === granularity ? "bg-teal-600 text-white shadow" : "text-gray-500 hover:text-gray-700"
                }`}
              >
                {g.charAt(0).toUpperCase() + g.slice(1)}
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Filters */}
      {attribute_options.length > 0 && (
        <AttributeFilterBar attributes={attribute_options} filters={filters} onChange={handleFilterChange} />
      )}

      {/* A) ARR Histogram */}
      <ARRHistogram data={data.arrHistogram} metricLabel={metricLabel} />

      {/* B & C) Mekko Charts */}
      <div className="space-y-2">
        <div className="flex items-center gap-4 flex-wrap">
          <span className="text-xs font-semibold text-gray-600 uppercase tracking-wide">Mekko Axes</span>
          <AxisSelector label="X-Axis" value={effectiveMekkoX} options={identifiers} onChange={handleMekkoXChange} />
          {identifiers.length > 1 && (
            <AxisSelector label="Y-Axis" value={effectiveMekkoY} options={identifiers} onChange={handleMekkoYChange} />
          )}
        </div>
        <MekkoChart
          data={data.mekkoARR}
          title={`${metricLabel} Distribution`}
          scaleFactor={scale_factor}
          valueType="arr"
          metricLabel={metricLabel}
        />
        <MekkoChart
          data={data.mekkoCount}
          title="Customer Count Distribution"
          scaleFactor={scale_factor}
          valueType="count"
          metricLabel={metricLabel}
        />
      </div>

      {/* D) Pie Charts */}
      <IdentifierPieCharts data={data.pieCharts} scaleFactor={scale_factor} metricLabel={metricLabel} />

      {/* E) Growth Histogram */}
      <GrowthHistogram data={data.growthHistogram} />

      {/* F & G) Grids */}
      <div className="space-y-2">
        <div className="flex items-center gap-4 flex-wrap">
          <span className="text-xs font-semibold text-gray-600 uppercase tracking-wide">Grid Axes</span>
          <AxisSelector label="X-Axis" value={effectiveGridX} options={identifiers} onChange={handleGridXChange} />
          {identifiers.length > 1 && (
            <AxisSelector label="Y-Axis" value={effectiveGridY} options={identifiers} onChange={handleGridYChange} />
          )}
        </div>

        {/* F) Growth Grid */}
        <TwoByTwoGrid
          data={data.growthGrid}
          title="YoY Growth by Segment"
          subtitle={periodSubtitle}
          formatMetric={formatGrowth}
          colorScale={growthColor}
        />

        {/* G) Retention Grids */}
        <RetentionGrids
          netRetention={data.netRetentionGrid}
          lossRetention={data.lossRetentionGrid}
          subtitle={periodSubtitle}
        />
      </div>
    </div>
  );
}
