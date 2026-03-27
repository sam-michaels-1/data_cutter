import type { WizardState, DataType } from "../../types/wizard";

interface Props {
  state: WizardState;
  dispatch: React.Dispatch<any>;
}

const OPTIONS: { value: DataType; label: string; desc: string; disabled: boolean }[] = [
  {
    value: "arr",
    label: "ARR (Annual Recurring Revenue)",
    desc: "Point-in-time snapshot. Duplicates are resolved by taking the max value per customer per period.",
    disabled: false,
  },
  {
    value: "revenue",
    label: "Revenue (Transactional)",
    desc: "Cumulative over a period. All rows are summed per customer per period.",
    disabled: true,
  },
];

export default function DataTypeStep({ state, dispatch }: Props) {
  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-semibold text-gray-800 mb-1">
          What type of data is this?
        </h2>
        <p className="text-gray-500 text-sm">
          This determines how duplicate entries for the same customer/period are
          handled.
        </p>
      </div>

      <div className="space-y-3">
        {OPTIONS.map((opt) => (
          <button
            key={opt.value}
            onClick={() =>
              !opt.disabled &&
              dispatch({ type: "SET_DATA_TYPE", dataType: opt.value })
            }
            disabled={opt.disabled}
            className={`w-full text-left border-2 rounded-xl p-4 transition
              ${state.dataType === opt.value ? "border-blue-500 bg-blue-50" : "border-gray-200 hover:border-gray-300"}
              ${opt.disabled ? "opacity-50 cursor-not-allowed" : "cursor-pointer"}
            `}
          >
            <div className="flex items-center justify-between">
              <span className="font-medium text-gray-800">{opt.label}</span>
              {opt.disabled && (
                <span className="text-xs bg-gray-200 text-gray-500 px-2 py-0.5 rounded-full">
                  Coming Soon
                </span>
              )}
              {state.dataType === opt.value && !opt.disabled && (
                <span className="text-blue-600 text-lg">{"\u2713"}</span>
              )}
            </div>
            <p className="text-sm text-gray-500 mt-1">{opt.desc}</p>
          </button>
        ))}
      </div>
    </div>
  );
}
