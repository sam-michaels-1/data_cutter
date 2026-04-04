import type { AttributeOption, Filters } from "../types/dashboard";
import MultiSelectDropdown from "./ui/MultiSelectDropdown";

interface Props {
  attributes: AttributeOption[];
  filters: Filters;
  onChange: (filters: Filters) => void;
}

export default function AttributeFilterBar({ attributes, filters, onChange }: Props) {
  if (attributes.length === 0) return null;

  const handleSingleChange = (attrName: string, value: string) => {
    const next = { ...filters };
    if (value === "") {
      delete next[attrName];
    } else {
      next[attrName] = value;
    }
    onChange(next);
  };

  const handleMultiChange = (attrName: string, selected: string[]) => {
    const attr = attributes.find(a => a.name === attrName);
    const next = { ...filters };
    // If all selected or none selected, remove filter (means "all")
    if (!attr || selected.length === attr.values.length) {
      delete next[attrName];
    } else {
      next[attrName] = selected;
    }
    onChange(next);
  };

  const hasActiveFilter = Object.keys(filters).length > 0;

  return (
    <div className="flex items-center gap-3 flex-wrap">
      {attributes.map(({ name, values, multiSelect }) => {
        if (multiSelect) {
          const selected = filters[name];
          const selectedArr = Array.isArray(selected) ? selected : values;
          return (
            <MultiSelectDropdown
              key={name}
              label={name}
              options={values}
              selected={selectedArr}
              onChange={(sel) => handleMultiChange(name, sel)}
            />
          );
        }

        return (
          <div key={name} className="flex items-center gap-1.5">
            <label className="text-xs text-gray-500 font-medium whitespace-nowrap">
              {name}
            </label>
            <select
              value={(filters[name] as string) ?? ""}
              onChange={(e) => handleSingleChange(name, e.target.value)}
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
        );
      })}
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
