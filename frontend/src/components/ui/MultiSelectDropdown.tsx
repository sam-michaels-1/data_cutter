import { useState, useRef, useEffect } from "react";

interface Props {
  label: string;
  options: string[];
  selected: string[];
  onChange: (selected: string[]) => void;
}

export default function MultiSelectDropdown({ label, options, selected, onChange }: Props) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false);
      }
    }
    if (open) document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [open]);

  const allSelected = selected.length === options.length;
  const noneSelected = selected.length === 0;

  let buttonText: string;
  if (allSelected) {
    buttonText = "All";
  } else if (noneSelected) {
    buttonText = "None";
  } else if (selected.length <= 2) {
    buttonText = selected.join(", ");
  } else {
    buttonText = `${selected.length} selected`;
  }

  const toggle = (val: string) => {
    if (selected.includes(val)) {
      onChange(selected.filter((s) => s !== val));
    } else {
      onChange([...selected, val]);
    }
  };

  return (
    <div ref={ref} className="relative">
      <div className="flex items-center gap-1.5">
        <span className="text-xs text-gray-500 font-medium whitespace-nowrap">{label}</span>
        <button
          onClick={() => setOpen(!open)}
          className="text-xs border border-gray-200 rounded-md px-2 py-1 bg-white text-gray-700 focus:outline-none focus:ring-1 focus:ring-teal-500 focus:border-teal-500 flex items-center gap-1 min-w-[80px]"
        >
          <span className="truncate max-w-[140px]">{buttonText}</span>
          <svg className="w-3 h-3 shrink-0 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
          </svg>
        </button>
      </div>

      {open && (
        <div className="absolute top-full left-0 mt-1 z-20 bg-white border border-gray-200 rounded-lg shadow-lg min-w-[180px] max-h-[280px] overflow-auto">
          <div className="flex items-center gap-2 px-3 py-2 border-b border-gray-100">
            <button
              onClick={() => onChange([...options])}
              className="text-[10px] font-medium text-teal-600 hover:text-teal-800 transition"
            >
              Select All
            </button>
            <span className="text-gray-300">|</span>
            <button
              onClick={() => onChange([])}
              className="text-[10px] font-medium text-teal-600 hover:text-teal-800 transition"
            >
              Deselect All
            </button>
          </div>
          <div className="py-1">
            {options.map((opt) => {
              const checked = selected.includes(opt);
              return (
                <label
                  key={opt}
                  className="flex items-center gap-2 px-3 py-1 hover:bg-gray-50 cursor-pointer"
                >
                  <input
                    type="checkbox"
                    checked={checked}
                    onChange={() => toggle(opt)}
                    className="rounded border-gray-300 text-teal-600 focus:ring-teal-500 h-3 w-3"
                  />
                  <span className="text-xs text-gray-700">{opt}</span>
                </label>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}
