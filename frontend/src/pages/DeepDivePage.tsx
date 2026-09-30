import { useMemo, useState } from "react";
import { Link, useNavigate, useParams } from "react-router-dom";
import {
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
} from "recharts";
import { useSession } from "../components/SessionProvider";
import { getCurrentWorkbook, getCurrentConfig } from "../api/client";
import { computeCustomers, computeCustomerDetail } from "../engine/customer_detail";
import { formatCurrency, formatPct } from "../utils/format";
import type { CustomerDetail, CustomerEventType } from "../types/dashboard";

const STATUS_STYLES: Record<string, string> = {
  Growth: "text-emerald-700 bg-emerald-100",
  Stable: "text-gray-600 bg-gray-100",
  Declining: "text-red-700 bg-red-100",
  New: "text-blue-700 bg-blue-100",
};

const EVENT_META: Record<CustomerEventType, { label: string; classes: string }> = {
  new: { label: "New logo", classes: "text-blue-700 bg-blue-100" },
  upsell: { label: "Upsell", classes: "text-emerald-700 bg-emerald-100" },
  downsell: { label: "Downsell", classes: "text-amber-700 bg-amber-100" },
  churn: { label: "Churned", classes: "text-red-700 bg-red-100" },
  reactivation: { label: "Returned", classes: "text-teal-700 bg-teal-100" },
  flat: { label: "Flat", classes: "text-gray-500 bg-gray-100" },
  inactive: { label: "—", classes: "text-gray-400 bg-gray-50" },
};

function NoDataCard({ onImport }: { onImport: () => void }) {
  return (
    <div className="flex items-center justify-center h-full min-h-[60vh]">
      <div className="text-center text-gray-500">
        <p className="text-lg font-medium">No data imported yet</p>
        <p className="text-sm mt-1">Please return to the Import tab to load your data.</p>
        <button
          onClick={onImport}
          className="mt-3 px-4 py-2 bg-teal-600 text-white rounded-lg text-sm hover:bg-teal-700 transition"
        >
          Go to Import
        </button>
      </div>
    </div>
  );
}

function StatCard({ label, value, sub, color = "text-gray-900" }: { label: string; value: string; sub?: string; color?: string }) {
  return (
    <div className="bg-white border border-gray-200 rounded-xl px-2.5 py-2 sm:px-3">
      <p className="text-xs font-medium text-gray-500 uppercase tracking-wide">{label}</p>
      <p className={`text-lg sm:text-xl font-bold mt-1 ${color}`}>{value}</p>
      {sub && <p className="text-xs text-gray-400 mt-0.5">{sub}</p>}
    </div>
  );
}

export default function DeepDivePage() {
  const { sessionId } = useSession();
  const navigate = useNavigate();
  const { customerName } = useParams<{ customerName?: string }>();
  const wb = getCurrentWorkbook();
  const config = getCurrentConfig();

  const [granularity, setGranularity] = useState<string | undefined>(undefined);
  const [query, setQuery] = useState("");

  const customers = useMemo(
    () => (wb && config ? computeCustomers(wb, config) : []),
    [wb, config]
  );
  const detail = useMemo<CustomerDetail | null>(
    () =>
      wb && config && customerName
        ? computeCustomerDetail(wb, config, customerName, granularity)
        : null,
    [wb, config, customerName, granularity]
  );

  const goImport = () => navigate("/import");

  if (!sessionId || !wb || !config) {
    return <NoDataCard onImport={goImport} />;
  }

  // ---------- Picker view (also used for unknown customer names) ----------
  if (!detail) {
    const q = query.trim().toLowerCase();
    const filtered = q
      ? customers.filter((c) => c.name.toLowerCase().includes(q))
      : customers;
    return (
      <div className="p-3 sm:p-4 space-y-3 max-w-[1600px]">
        <div>
          <h1 className="text-xl font-bold">Customer Deep Dive</h1>
          <p className="text-sm text-gray-500">Select a customer to see their full history</p>
        </div>

        {customerName && (
          <div className="bg-amber-50 border border-amber-200 text-amber-800 text-sm rounded-lg px-3 py-2">
            Customer &ldquo;{customerName}&rdquo; was not found in this workbook. Pick a customer below.
          </div>
        )}

        <div className="bg-white border border-gray-200 rounded-xl p-3">
          <input
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search customers..."
            className="w-full mb-2 px-3 py-2 text-sm border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-teal-500"
            autoFocus
          />
          <div className="overflow-auto max-h-[65vh] divide-y divide-gray-100">
            {filtered.map((c) => (
              <Link
                key={c.name}
                to={`/deep-dive/${encodeURIComponent(c.name)}`}
                className="flex items-center justify-between px-2 py-2 hover:bg-teal-50 rounded-md group"
              >
                <span className="text-sm font-medium text-gray-800 group-hover:text-teal-700 truncate">
                  {c.name}
                </span>
                <span className="text-sm font-mono text-gray-500 ml-4 shrink-0">
                  {formatCurrency(c.current_arr, config.scale_factor)}
                </span>
              </Link>
            ))}
            {filtered.length === 0 && (
              <p className="px-2 py-6 text-sm text-gray-400 text-center">
                No customers match &ldquo;{query}&rdquo;
              </p>
            )}
          </div>
        </div>
      </div>
    );
  }

  // ---------- Detail view ----------
  const metricLabel = detail.data_type === "revenue" ? "Revenue" : "ARR";
  const sf = detail.scale_factor;
  const chartData = detail.timeline.map((t) => ({ period: t.period_label, arr: t.arr }));
  const attrEntries = Object.entries(detail.attributes).filter(([, v]) => v);

  const changeColor = (v: number | null) =>
    v == null ? "text-gray-900" : v > 0 ? "text-emerald-600" : v < 0 ? "text-red-500" : "text-gray-900";

  return (
    <div className="p-3 sm:p-4 space-y-3 max-w-[1600px]">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-start sm:justify-between gap-2">
        <div>
          <Link to="/deep-dive" className="text-xs text-teal-600 hover:underline">
            &larr; All customers
          </Link>
          <h1 className="text-xl font-bold text-gray-900">{detail.name}</h1>
          <div className="flex flex-wrap items-center gap-1.5 mt-1.5">
            {attrEntries.map(([k, v]) => (
              <span
                key={k}
                className="inline-block px-2 py-0.5 rounded-full text-xs bg-gray-100 text-gray-600"
              >
                <span className="text-gray-400">{k}:</span> {v}
              </span>
            ))}
            {detail.cohort && (
              <span className="inline-block px-2 py-0.5 rounded-full text-xs bg-teal-50 text-teal-700">
                Cohort: {detail.cohort}
              </span>
            )}
            <span
              className={`inline-block px-2 py-0.5 rounded-full text-xs font-medium ${
                STATUS_STYLES[detail.status] || STATUS_STYLES.Stable
              }`}
            >
              {detail.status}
            </span>
          </div>
        </div>

        {detail.available_granularities.length > 1 && (
          <div className="flex gap-1 bg-gray-100 rounded-lg p-0.5 self-start">
            {detail.available_granularities.map((g) => (
              <button
                key={g}
                onClick={() => setGranularity(g)}
                className={`px-3 py-1 rounded-md text-xs font-medium transition ${
                  g === detail.granularity
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

      {/* Stat cards */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <StatCard label={`Current ${metricLabel}`} value={formatCurrency(detail.current_arr, sf)} />
        <StatCard
          label={`First ${metricLabel}`}
          value={formatCurrency(detail.first_arr, sf)}
          sub={detail.first_period_label}
        />
        <StatCard
          label={`Peak ${metricLabel}`}
          value={formatCurrency(detail.peak_arr, sf)}
          sub={detail.peak_period_label}
        />
        <StatCard label="Lifetime Total" value={formatCurrency(detail.lifetime_total, sf)} />
        <StatCard label="% of Total" value={(detail.pct_of_total * 100).toFixed(1) + "%"} />
        <StatCard
          label="Total Change"
          value={detail.total_change != null ? formatCurrency(detail.total_change, sf) : "N/A"}
          sub={formatPct(detail.total_change_pct)}
          color={changeColor(detail.total_change_pct)}
        />
        <StatCard
          label="YoY Change"
          value={formatPct(detail.yoy_change_pct)}
          color={changeColor(detail.yoy_change_pct)}
        />
        <StatCard
          label="CAGR"
          value={formatPct(detail.cagr)}
          color={changeColor(detail.cagr)}
        />
      </div>

      {/* Timeline chart */}
      <div className="bg-white border border-gray-200 rounded-xl p-4">
        <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-3">
          {metricLabel} Over Time
        </h3>
        <div className="h-64">
          <ResponsiveContainer width="100%" height="100%">
            <AreaChart data={chartData} margin={{ top: 4, right: 8, left: 0, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
              <XAxis dataKey="period" tick={{ fontSize: 11 }} interval="preserveStartEnd" />
              <YAxis
                tick={{ fontSize: 11 }}
                tickFormatter={(v: number) => formatCurrency(v, sf)}
                width={70}
              />
              <Tooltip
                formatter={(value) => [formatCurrency(Number(value), sf), metricLabel]}
              />
              <Area
                type="monotone"
                dataKey="arr"
                stroke="#14B8A6"
                fill="#14B8A6"
                fillOpacity={0.15}
                strokeWidth={2}
              />
            </AreaChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-3">
        {/* Period-over-period table */}
        <div className="lg:col-span-2 bg-white border border-gray-200 rounded-xl p-3">
          <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-2">
            Period over Period
          </h3>
          <div className="overflow-auto max-h-[420px]">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-xs text-gray-500 uppercase tracking-wide border-b border-gray-200">
                  <th className="sticky top-0 z-10 text-left py-2 pr-4 bg-white">Period</th>
                  <th className="sticky top-0 z-10 text-right py-2 pr-4 bg-white">{metricLabel}</th>
                  <th className="sticky top-0 z-10 text-right py-2 pr-4 bg-white">Change</th>
                  <th className="sticky top-0 z-10 text-right py-2 pr-4 bg-white">% Change</th>
                  <th className="sticky top-0 z-10 text-center py-2 bg-white">Event</th>
                </tr>
              </thead>
              <tbody>
                {[...detail.timeline].reverse().map((t) => {
                  const highlight =
                    t.event === "new"
                      ? "bg-blue-50/60"
                      : t.event === "churn"
                      ? "bg-red-50/60"
                      : t.event === "reactivation"
                      ? "bg-teal-50/60"
                      : "";
                  const meta = EVENT_META[t.event];
                  return (
                    <tr key={t.period_label} className={`border-b border-gray-100/50 ${highlight}`}>
                      <td className="py-2 pr-4 text-gray-700 whitespace-nowrap">{t.period_label}</td>
                      <td className="py-2 pr-4 text-right font-mono text-gray-800">
                        {t.arr > 0 ? formatCurrency(t.arr, sf) : "—"}
                      </td>
                      <td className={`py-2 pr-4 text-right font-mono ${changeColor(t.change)}`}>
                        {t.change == null
                          ? "—"
                          : t.change === 0
                          ? "0"
                          : formatCurrency(t.change, sf)}
                      </td>
                      <td className={`py-2 pr-4 text-right font-mono ${changeColor(t.change_pct)}`}>
                        {t.change_pct == null
                          ? "—"
                          : t.change_pct < 0
                          ? `(${(Math.abs(t.change_pct) * 100).toFixed(1)}%)`
                          : `+${(t.change_pct * 100).toFixed(1)}%`}
                      </td>
                      <td className="py-2 text-center">
                        <span
                          className={`inline-block px-2 py-0.5 rounded-full text-xs font-medium ${meta.classes}`}
                        >
                          {meta.label}
                        </span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* Movement summary + attribute history */}
        <div className="space-y-3">
          <div className="bg-white border border-gray-200 rounded-xl p-3">
            <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-2">
              Movements
            </h3>
            {detail.movements.length === 0 ? (
              <p className="text-sm text-gray-400">No movements recorded.</p>
            ) : (
              <ul className="space-y-1.5 overflow-auto max-h-[300px]">
                {[...detail.movements].reverse().map((m, i) => {
                  const meta = EVENT_META[m.type];
                  return (
                    <li key={i} className="flex items-center gap-2 text-sm">
                      <span
                        className={`inline-block px-2 py-0.5 rounded-full text-xs font-medium shrink-0 ${meta.classes}`}
                      >
                        {meta.label}
                      </span>
                      <span className="text-gray-500 text-xs whitespace-nowrap">{m.period_label}</span>
                      <span className={`ml-auto font-mono text-xs ${changeColor(m.amount)}`}>
                        {m.type === "new" || m.type === "reactivation"
                          ? formatCurrency(m.amount, sf)
                          : m.amount < 0
                          ? `(${formatCurrency(Math.abs(m.amount), sf)})`
                          : `+${formatCurrency(m.amount, sf)}`}
                      </span>
                    </li>
                  );
                })}
              </ul>
            )}
          </div>

          {detail.attribute_transitions.length > 0 && (
            <div className="bg-white border border-gray-200 rounded-xl p-3">
              <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-2">
                Attribute History
              </h3>
              <ul className="space-y-1.5">
                {detail.attribute_transitions.map((t, i) => (
                  <li key={i} className="text-sm">
                    <span className="text-gray-500 text-xs">{t.period_label}</span>
                    <span className="mx-1.5 text-gray-400 text-xs">({t.date})</span>
                    <span className="font-medium text-gray-700">{t.attribute}:</span>{" "}
                    <span className="text-gray-500">{t.from || "—"}</span>
                    <span className="mx-1 text-gray-400">&rarr;</span>
                    <span className="text-teal-700 font-medium">{t.to || "—"}</span>
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
