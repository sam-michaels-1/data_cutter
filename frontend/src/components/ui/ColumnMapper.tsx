import type { ColumnInfo, ColumnMapping } from "../../types/wizard";

interface Props {
  columns: ColumnInfo[];
  mapping: ColumnMapping;
  onChange: (mapping: ColumnMapping) => void;
}

const ROLES = [
  { key: "date_col" as const, label: "Date / Period" },
  { key: "customer_id_col" as const, label: "Customer ID" },
  { key: "arr_col" as const, label: "ARR / Revenue" },
];

export default function ColumnMapper({ columns, mapping, onChange }: Props) {
  return (
    <div className="bg-white border border-gray-200 rounded-lg overflow-hidden">
      <table className="w-full text-sm">
        <thead className="bg-gray-50">
          <tr>
            <th className="px-4 py-2 text-left font-medium text-gray-600">
              Column
            </th>
            <th className="px-4 py-2 text-left font-medium text-gray-600">
              Header
            </th>
            <th className="px-4 py-2 text-left font-medium text-gray-600">
              Sample Values
            </th>
          </tr>
        </thead>
        <tbody>
          {columns.map((col) => {
            const assignedRole = ROLES.find(
              (r) => mapping[r.key] === col.letter
            );
            return (
              <tr
                key={col.letter}
                className={`border-t ${assignedRole ? "bg-blue-50" : ""}`}
              >
                <td className="px-4 py-2 font-mono font-bold text-gray-700">
                  {col.letter}
                  {assignedRole && (
                    <span className="ml-2 text-xs bg-blue-600 text-white px-2 py-0.5 rounded-full">
                      {assignedRole.label}
                    </span>
                  )}
                </td>
                <td className="px-4 py-2 text-gray-800">{col.header}</td>
                <td className="px-4 py-2 text-gray-500 text-xs truncate max-w-48">
                  {col.sample_values.slice(0, 3).join(", ")}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>

      <div className="bg-gray-50 px-4 py-3 border-t">
        <p className="text-xs text-gray-500 mb-2 font-medium">
          Assign column roles:
        </p>
        <div className="flex flex-wrap gap-4">
          {ROLES.map((role) => (
            <div key={role.key} className="flex items-center gap-2">
              <label className="text-sm font-medium text-gray-700 min-w-28">
                {role.label}:
              </label>
              <select
                value={mapping[role.key]}
                onChange={(e) =>
                  onChange({ ...mapping, [role.key]: e.target.value })
                }
                className="border border-gray-300 rounded px-2 py-1 text-sm bg-white"
              >
                {columns.map((col) => (
                  <option key={col.letter} value={col.letter}>
                    {col.letter} — {col.header}
                  </option>
                ))}
              </select>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
