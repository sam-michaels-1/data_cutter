import type { AttributeOption } from "../types/dashboard";

interface Props {
  attributes: AttributeOption[];
  filters: Record<string, string>;
  onChange: (filters: Record<string, string>) => void;
}

export default function AttributeFilterBar({ attributes, filters, onChange }: Props) {
  if (attributes.length === 0) return null;

  const handleChange = (attrName: string, value: string) => {
    const next = { ...filters };
    if (value === "") {
      delete next[attrName];
    } else {
      next[attrName] = value;
    }
    onChange(next);
  };

  const hasActiveFilter = Object.keys(filters).length > 0;

  return (
    <div className="flex items-center gap-3 flex-wrap">
      {attributes.map(({ name, values }) => (
        <div key={name} className="flex items-center gap-1.5">
          <label className="text-xs text-gray-500 font-medium whitespace-nowrap">
            {name}
          </label>
          <select
            value={filters[name] ?? ""}
            onChange={(e) => handleChange(name, e.target.value)}
            className="text-xs border border-gray-200 rounded-md px-2 py-1 bg-white text-gray-700 focus:outline-none focus:ring-1 focus:ring-teal-500 focus:border-teal-500"
          >
            <option value="">All</option>
            {values.map((v) => (
              <option key={v} value={v}>
                {v}
              </option>
            ))}
          </select>
        </div>
      ))}
      {hasActiveFilter && (
        <button
          onClick={() => onChange({})}
          className="text-xs text-teal-600 hover:text-teal-800 font-medium transition"
        >
          Clear filters
        </button>
      )}
    </div>
  );
}
