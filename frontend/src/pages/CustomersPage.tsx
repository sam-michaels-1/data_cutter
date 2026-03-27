import { LineChart, Line, ResponsiveContainer } from "recharts";
import { useSession } from "../components/SessionProvider";
import { useDashboard } from "../hooks/useDashboard";
import { formatCurrency } from "../utils/format";

const STATUS_STYLES: Record<string, string> = {
  Growth: "text-emerald-700 bg-emerald-100",
  Stable: "text-gray-600 bg-gray-100",
  Declining: "text-red-700 bg-red-100",
  New: "text-blue-700 bg-blue-100",
};

export default function CustomersPage() {
  const { sessionId } = useSession();
  const { data, loading, error, refetch } = useDashboard(sessionId);

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
          <p className="text-sm">Loading customer data...</p>
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

  const customers = data.overview.top_customers;
  const periods = data.overview.periods;
  const scaleFactor = data.scale_factor;
  const attrKeys = customers.length > 0 ? Object.keys(customers[0].attributes) : [];

  return (
    <div className="p-6 space-y-4 max-w-[1400px]">
      <div>
        <h1 className="text-xl font-bold">Customer Ranking</h1>
        <p className="text-sm text-gray-500">
          Top customers ranked by latest period ARR
        </p>
      </div>

      <div className="bg-white border border-gray-200 rounded-xl p-4">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-xs text-gray-500 uppercase tracking-wide border-b border-gray-200">
                <th className="text-left py-2 pr-2 w-8">#</th>
                <th className="text-left py-2 pr-4">Customer</th>
                {attrKeys.map((k) => (
                  <th key={k} className="text-left py-2 pr-4">{k}</th>
                ))}
                <th className="text-left py-2 pr-4">Cohort</th>
                <th className="text-right py-2 pr-4">Current ARR</th>
                <th className="text-right py-2 pr-4">Change</th>
                <th className="text-right py-2 pr-4">% of Total</th>
                <th className="text-center py-2 pr-4 w-20">Trend</th>
                <th className="text-center py-2 pr-4">Status</th>
                {periods.map((p) => (
                  <th key={p} className="text-right py-2 pr-3 whitespace-nowrap">
                    {p}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {customers.map((c) => {
                const sparkData = c.trend.map((v, i) => ({ i, v }));
                return (
                  <tr key={c.rank} className="border-b border-gray-100/50">
                    <td className="py-2.5 pr-2 text-gray-500">{c.rank}</td>
                    <td className="py-2.5 pr-4 font-medium text-gray-800 truncate max-w-[160px]">
                      {c.name}
                    </td>
                    {attrKeys.map((k) => (
                      <td key={k} className="py-2.5 pr-4 text-gray-600 whitespace-nowrap">
                        {c.attributes[k] || "—"}
                      </td>
                    ))}
                    <td className="py-2.5 pr-4 text-gray-600 whitespace-nowrap">{c.cohort || "—"}</td>
                    <td className="py-2.5 pr-4 text-right font-mono text-gray-800">
                      {formatCurrency(c.arr, scaleFactor)}
                    </td>
                    <td
                      className={`py-2.5 pr-4 text-right font-mono ${
                        c.change_pct == null
                          ? "text-gray-400"
                          : c.change_pct > 0
                          ? "text-emerald-500"
                          : c.change_pct < 0
                          ? "text-red-400"
                          : "text-gray-400"
                      }`}
                    >
                      {c.change_pct == null
                        ? "N/A"
                        : `${c.change_pct > 0 ? "+" : ""}${(c.change_pct * 100).toFixed(1)}%`}
                    </td>
                    <td className="py-2.5 pr-4 text-right font-mono text-gray-600">
                      {(c.pct_of_total * 100).toFixed(1)}%
                    </td>
                    <td className="py-2.5 pr-4">
                      <div className="w-20 h-7 mx-auto">
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={sparkData}>
                            <Line
                              type="monotone"
                              dataKey="v"
                              stroke="#14B8A6"
                              strokeWidth={1.5}
                              dot={false}
                            />
                          </LineChart>
                        </ResponsiveContainer>
                      </div>
                    </td>
                    <td className="py-2.5 pr-4 text-center">
                      <span
                        className={`inline-block px-2 py-0.5 rounded-full text-xs font-medium ${
                          STATUS_STYLES[c.status] || STATUS_STYLES.Stable
                        }`}
                      >
                        {c.status}
                      </span>
                    </td>
                    {c.trend.map((val, i) => (
                      <td
                        key={i}
                        className="py-2.5 pr-3 text-right font-mono text-gray-600 whitespace-nowrap"
                      >
                        {val > 0 ? formatCurrency(val, scaleFactor) : "—"}
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
