import type { WizardState, DataFrequency } from "../../types/wizard";

interface Props {
  state: WizardState;
  dispatch: React.Dispatch<any>;
}

const OPTIONS: { value: DataFrequency; label: string; desc: string }[] = [
  {
    value: "monthly",
    label: "Monthly",
    desc: "Data contains one row per customer per month (e.g., dates spaced ~30 days apart).",
  },
  {
    value: "quarterly",
    label: "Quarterly",
    desc: "Data contains one row per customer per quarter (e.g., dates spaced ~90 days apart).",
  },
];

export default function FrequencyStep({ state, dispatch }: Props) {
  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-semibold text-gray-800 mb-1">
          What is the frequency of your data?
        </h2>
        <p className="text-gray-500 text-sm">
          This determines how dates are interpreted in the analysis.
          {state.detectedFrequency && (
            <>
              {" "}We detected <span className="font-medium text-gray-700">{state.detectedFrequency}</span> data.
            </>
          )}
        </p>
      </div>

      <div className="space-y-3">
        {OPTIONS.map((opt) => (
          <button
            key={opt.value}
            onClick={() =>
              dispatch({ type: "SET_DATA_FREQUENCY", dataFrequency: opt.value })
            }
            className={`w-full text-left border-2 rounded-xl p-4 transition cursor-pointer
              ${state.dataFrequency === opt.value ? "border-blue-500 bg-blue-50" : "border-gray-200 hover:border-gray-300"}
            `}
          >
            <div className="flex items-center justify-between">
              <span className="font-medium text-gray-800">{opt.label}</span>
              {state.dataFrequency === opt.value && (
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
