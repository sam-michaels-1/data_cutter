import type { WizardState } from "../../types/wizard";

interface Props {
  state: WizardState;
  dispatch: React.Dispatch<any>;
}

export default function IdentifiersStep({ state, dispatch }: Props) {
  const { detectedAttributes, selectedAttributes } = state;

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-semibold text-gray-800 mb-1">
          Key Identifiers
        </h2>
        <p className="text-gray-500 text-sm">
          Select which columns to use as analysis dimensions (filters, breakouts,
          cohort cuts). You can rename them to friendlier labels.
        </p>
      </div>

      {detectedAttributes.length === 0 ? (
        <p className="text-gray-400 text-sm italic">
          No attribute columns were detected. You can proceed without
          identifiers.
        </p>
      ) : (
        <div className="space-y-2">
          {detectedAttributes.map((attr) => {
            const selected = selectedAttributes.find(
              (a) => a.letter === attr.letter
            );
            return (
              <div
                key={attr.letter}
                className={`flex items-center gap-4 border rounded-lg p-3 transition
                  ${selected ? "border-blue-300 bg-blue-50" : "border-gray-200"}
                `}
              >
                {/* Toggle */}
                <button
                  onClick={() =>
                    dispatch({ type: "TOGGLE_ATTRIBUTE", attr })
                  }
                  className={`relative inline-flex h-6 w-11 flex-shrink-0 items-center rounded-full transition-colors
                    ${selected ? "bg-blue-600" : "bg-gray-300"}
                  `}
                >
                  <span
                    className={`inline-block h-5 w-5 rounded-full bg-white shadow transition-transform
                      ${selected ? "translate-x-5" : "translate-x-0.5"}
                    `}
                  />
                </button>

                {/* Column info */}
                <span className="font-mono text-sm text-gray-500 w-8">
                  {attr.letter}
                </span>

                {/* Display name (editable if selected) */}
                {selected ? (
                  <input
                    type="text"
                    value={selected.display_name}
                    onChange={(e) =>
                      dispatch({
                        type: "RENAME_ATTRIBUTE",
                        letter: attr.letter,
                        newName: e.target.value,
                      })
                    }
                    className="flex-1 border border-blue-200 rounded px-2 py-1 text-sm bg-white"
                  />
                ) : (
                  <span className="flex-1 text-sm text-gray-500">
                    {attr.header}
                  </span>
                )}
              </div>
            );
          })}
        </div>
      )}

      {selectedAttributes.length > 0 && (
        <p className="text-xs text-gray-400">
          {selectedAttributes.length} identifier
          {selectedAttributes.length !== 1 && "s"} selected. The first one will
          be used to auto-generate filter breakouts.
        </p>
      )}
    </div>
  );
}
