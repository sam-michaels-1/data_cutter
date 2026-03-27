import { LineChart, Line, ResponsiveContainer } from "recharts";
import type { TopCustomer } from "../../types/dashboard";
import { formatCurrency } from "../../utils/format";

interface Props {
  customers: TopCustomer[];
  scaleFactor: number;
}

const STATUS_STYLES: Record<string, string> = {
  Growth: "text-emerald-700 bg-emerald-100",
  Stable: "text-gray-600 bg-gray-100",
  Declining: "text-red-700 bg-red-100",
  New: "text-blue-700 bg-blue-100",
};

export default function TopCustomersTable({ customers, scaleFactor }: Props) {
  // Derive attribute column names from the first customer
  const attrKeys = customers.length > 0 ? Object.keys(customers[0].attributes) : [];

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4">
      <h3 className="text-sm font-semibold text-gray-700 uppercase tracking-wide mb-3">
        Top 10 Customers by ARR
      </h3>
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
              <th className="text-right py-2 pr-4">ARR</th>
              <th className="text-right py-2 pr-4">Change</th>
              <th className="text-right py-2 pr-4">% of Total</th>
              <th className="text-center py-2 pr-4 w-16">Trend</th>
              <th className="text-center py-2">Status</th>
            </tr>
          </thead>
          <tbody>
            {customers.map((c) => {
              const sparkData = c.trend.map((v, i) => ({ i, v }));
              const changePct = c.change_pct;
              return (
                <tr
                  key={c.rank}
                  className="border-b border-gray-100 hover:bg-gray-50"
                >
                  <td className="py-2 pr-2 text-gray-500">{c.rank}</td>
                  <td className="py-2 pr-4 font-medium text-gray-800 truncate max-w-[160px]">
                    {c.name}
                  </td>
                  {attrKeys.map((k) => (
                    <td key={k} className="py-2 pr-4 text-gray-600 whitespace-nowrap">
                      {c.attributes[k] || "—"}
                    </td>
                  ))}
                  <td className="py-2 pr-4 text-gray-600 whitespace-nowrap">{c.cohort || "—"}</td>
                  <td className="py-2 pr-4 text-right font-mono text-gray-800">
                    {formatCurrency(c.arr, scaleFactor)}
                  </td>
                  <td
                    className={`py-2 pr-4 text-right font-mono ${
                      changePct == null
                        ? "text-gray-400"
                        : changePct > 0
                        ? "text-emerald-600"
                        : changePct < 0
                        ? "text-red-500"
                        : "text-gray-400"
                    }`}
                  >
                    {changePct == null
                      ? "N/A"
                      : `${changePct > 0 ? "+" : ""}${(changePct * 100).toFixed(1)}%`}
                  </td>
                  <td className="py-2 pr-4 text-right font-mono text-gray-600">
                    {(c.pct_of_total * 100).toFixed(1)}%
                  </td>
                  <td className="py-2 pr-4">
                    <div className="w-16 h-6 mx-auto">
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
                  <td className="py-2 text-center">
                    <span
                      className={`inline-block px-2 py-0.5 rounded-full text-xs font-medium ${
                        STATUS_STYLES[c.status] || STATUS_STYLES.Stable
                      }`}
                    >
                      {c.status}
                    </span>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
