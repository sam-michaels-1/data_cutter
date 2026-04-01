import type { WizardState, Granularity } from "../../types/wizard";

interface Props {
  state: WizardState;
  dispatch: React.Dispatch<any>;
}

const GRANULARITIES: { value: Granularity; label: string; desc: string }[] = [
  {
    value: "monthly",
    label: "Monthly",
    desc: "Monthly retention, cohort, and clean data tabs",
  },
  {
    value: "quarterly",
    label: "Quarterly",
    desc: "Quarterly retention, cohort, and clean data tabs",
  },
  {
    value: "annual",
    label: "Annual",
    desc: "Annual retention, cohort, clean data, and top customers",
  },
];

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

export default function GranularityStep({ state, dispatch }: Props) {
  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-semibold text-gray-800 mb-1">
          Output Granularity
        </h2>
        <p className="text-gray-500 text-sm">
          Select which time periods you want in the output. Pick one or more.
        </p>
      </div>

      <div className="space-y-3">
        {GRANULARITIES.map((g) => {
          const selected = state.outputGranularities.includes(g.value);
          const disabled =
            g.value === "monthly" && state.dataFrequency === "quarterly";
          return (
            <button
              key={g.value}
              disabled={disabled}
              onClick={() =>
                dispatch({ type: "TOGGLE_GRANULARITY", granularity: g.value })
              }
              className={`w-full text-left border-2 rounded-xl p-4 transition
                ${disabled ? "border-gray-100 bg-gray-50 opacity-50 cursor-not-allowed" : "cursor-pointer"}
                ${!disabled && selected ? "border-blue-500 bg-blue-50" : ""}
                ${!disabled && !selected ? "border-gray-200 hover:border-gray-300" : ""}
              `}
            >
              <div className="flex items-center justify-between">
                <span className={`font-medium ${disabled ? "text-gray-400" : "text-gray-800"}`}>{g.label}</span>
                {selected && !disabled && (
                  <span className="text-blue-600 text-lg">{"\u2713"}</span>
                )}
              </div>
              <p className={`text-sm mt-1 ${disabled ? "text-gray-400" : "text-gray-500"}`}>
                {disabled ? "Not available for quarterly data" : g.desc}
              </p>
            </button>
          );
        })}
      </div>

      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">
          Fiscal Year End Month
        </label>
        <select
          value={state.fiscalYearEndMonth}
          onChange={(e) =>
            dispatch({
              type: "SET_FISCAL_MONTH",
              month: Number(e.target.value),
            })
          }
          className="w-full max-w-xs border border-gray-300 rounded-lg px-3 py-2 bg-white"
        >
          {MONTHS.map((m, i) => (
            <option key={i + 1} value={i + 1}>
              {m}
            </option>
          ))}
        </select>
      </div>
    </div>
  );
}
