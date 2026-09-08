import { useState, useCallback } from "react";
import { useNavigate } from "react-router-dom";
import { useSession } from "../components/SessionProvider";
import { useSummaryData } from "../hooks/useSummaryData";
import AttributeFilterBar from "../components/AttributeFilterBar";
import { retentionColor } from "../components/histograms/colorScales";
import type { SummarySection } from "../engine/summary_compute";
import { formatCurrency } from "../utils/format";
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

const RETENTION_KEYS = new Set<SummarySection["key"]>(["gross", "net", "logo"]);

export default function SummaryPage() {
  const { sessionId } = useSession();
  const navigate = useNavigate();
  const { data, loading, error, refetch } = useSummaryData(sessionId);
  const [filters, setFilters] = useState<Filters>({});
  const [identifier, setIdentifier] = useState("");

  const currentGran = data?.granularity;

  const doRefetch = useCallback((
    gran?: string,
    opts?: { filters?: Filters; identifier?: string }
  ) => {
    refetch(gran || currentGran, {
      filters: opts?.filters ?? filters,
      identifier: opts?.identifier ?? identifier,
    });
  }, [refetch, currentGran, filters, identifier]);

  const handleGranularityChange = (g: string) => doRefetch(g);
  const handleFilterChange = (f: Filters) => { setFilters(f); doRefetch(undefined, { filters: f }); };
  const handleIdentifierChange = (v: string) => { setIdentifier(v); doRefetch(undefined, { identifier: v }); };

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
        <div className="text-center text-gray-500">
          <div className="animate-spin h-8 w-8 border-2 border-teal-500 border-t-transparent rounded-full mx-auto mb-3" />
          <p className="text-sm">Computing summary...</p>
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

  const { identifiers, available_granularities, granularity, scale_factor, data_type, attribute_options, sections, columns } = data;
  const metricLabel = data_type === "revenue" ? "Revenue" : "ARR";
  const effectiveIdentifier = identifiers.includes(identifier) ? identifier : data.identifier;
  const segmentColumns = columns.slice(1);

  function formatCell(section: SummarySection, v: number | null): string {
    if (v == null) return "n.a.";
    if (section.format === "pct") return `${(v * 100).toFixed(0)}%`;
    if (section.format === "currency") return formatCurrency(v, scale_factor);
    return v.toLocaleString();
  }

  return (
    <div className="p-3 sm:p-4 space-y-3 max-w-[1600px]">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
        <div>
          <h1 className="text-xl font-bold text-gray-900">Summary</h1>
          <p className="hidden sm:block text-sm text-gray-500">
            Period metrics split by {effectiveIdentifier} — {metricLabel} basis
          </p>
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

      {/* Filters + column selector */}
      <div className="flex items-center gap-2 sm:gap-4 flex-wrap">
        <AxisSelector label="Columns" value={effectiveIdentifier} options={identifiers} onChange={handleIdentifierChange} />
        {attribute_options.length > 0 && (
          <AttributeFilterBar attributes={attribute_options} filters={filters} onChange={handleFilterChange} />
        )}
      </div>

      {/* Summary table */}
      <div className="bg-white border border-gray-200 rounded-xl p-3">
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr className="text-xs text-gray-500 uppercase tracking-wide border-b border-gray-200">
                <th className="text-left py-2 pr-3 font-medium whitespace-nowrap">Period</th>
                <th className="text-right py-2 px-2 font-medium whitespace-nowrap bg-gray-50">All</th>
                {segmentColumns.map(v => (
                  <th key={v} className="text-right py-2 px-2 font-medium whitespace-nowrap">{v}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sections.map((section, si) => (
                section.rows.length === 0 ? null : [
                  <tr key={`${section.key}-hdr`}>
                    <td
                      colSpan={columns.length + 1}
                      className={`text-[11px] font-semibold text-gray-700 uppercase tracking-wide pb-1 ${si === 0 ? "" : "pt-3"}`}
                    >
                      {section.label}
                    </td>
                  </tr>,
                  ...section.rows.map(row => (
                    <tr key={`${section.key}-${row.period}`} className="border-t border-gray-100 hover:bg-gray-50">
                      <td className="py-1.5 pr-3 text-gray-600 font-medium whitespace-nowrap">{row.period}</td>
                      {row.values.map((v, ci) => {
                        const colors = RETENTION_KEYS.has(section.key) && v != null ? retentionColor(v) : null;
                        return (
                          <td
                            key={ci}
                            className={`text-right font-mono tabular-nums px-2 py-1.5 whitespace-nowrap ${ci === 0 ? "bg-gray-50 font-medium" : "text-gray-700"}`}
                            style={colors ? { backgroundColor: colors.bg, color: colors.text } : undefined}
                          >
                            {formatCell(section, v)}
                          </td>
                        );
                      })}
                    </tr>
                  )),
                ]
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <p className="text-[10px] text-gray-400">
        Gross Retention = (BoP + Churn + Downsell) / BoP; Net Retention adds Upsell; Logo Retention = (BoP customers − churned customers) / BoP customers. BoP = the prior-year period.
      </p>
    </div>
  );
}
