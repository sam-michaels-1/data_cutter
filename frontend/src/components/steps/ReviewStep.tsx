import { useCallback } from "react";
import type { WizardState } from "../../types/wizard";
import { generate, getDownloadBlob } from "../../api/client";

interface Props {
  state: WizardState;
  dispatch: React.Dispatch<any>;
  onViewDashboard?: () => void;
}

const MONTHS = [
  "", "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

export default function ReviewStep({ state, dispatch, onViewDashboard }: Props) {
  const handleGenerate = useCallback(async () => {
    if (!state.sessionId || !state.selectedSheet || !state.confirmedMapping)
      return;
    dispatch({ type: "GENERATE_START" });
    try {
      const res = await generate({
        session_id: state.sessionId,
        sheet_name: state.selectedSheet,
        data_type: state.dataType,
        data_frequency: state.dataFrequency,
        column_mapping: state.confirmedMapping,
        attributes: state.selectedAttributes,
        output_granularities: state.outputGranularities,
        fiscal_year_end_month: state.fiscalYearEndMonth,
      });
      dispatch({ type: "GENERATE_SUCCESS", downloadId: res.download_id });
    } catch (err: any) {
      dispatch({
        type: "SET_ERROR",
        error: err instanceof Error ? err.message : "Generation failed",
      });
    }
  }, [state, dispatch]);

  const handleDownload = useCallback(() => {
    const blob = getDownloadBlob();
    if (!blob) return;
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "data-pack-output.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, []);

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-semibold text-gray-800 mb-1">
          Review & Generate
        </h2>
        <p className="text-gray-500 text-sm">
          Confirm your settings, then generate the analysis workbook.
        </p>
      </div>

      {/* Summary */}
      <div className="bg-gray-50 border border-gray-200 rounded-xl p-5 space-y-3 text-sm">
        <Row label="File" value={state.filename || "-"} />
        <Row label="Sheet" value={state.selectedSheet || "-"} />
        <Row label="Data Frequency" value={state.dataFrequency ? state.dataFrequency.charAt(0).toUpperCase() + state.dataFrequency.slice(1) : "-"} />
        <Row label="Data Type" value={state.dataType.toUpperCase()} />
        <Row
          label="Columns"
          value={`Date=${state.confirmedMapping?.date_col}, Customer=${state.confirmedMapping?.customer_id_col}, ARR=${state.confirmedMapping?.arr_col}`}
        />
        <Row
          label="Output Granularity"
          value={
            state.outputGranularities
              .map((g) => g.charAt(0).toUpperCase() + g.slice(1))
              .join(", ") || "None selected"
          }
        />
        <Row
          label="Fiscal Year End"
          value={MONTHS[state.fiscalYearEndMonth]}
        />
        <Row
          label="Identifiers"
          value={
            state.selectedAttributes.length > 0
              ? state.selectedAttributes.map((a) => a.display_name).join(", ")
              : "None"
          }
        />
        <Row
          label="Data Rows"
          value={state.rowCount.toLocaleString()}
        />
        <Row
          label="Scale Factor"
          value={state.scaleFactor.toLocaleString()}
        />
      </div>

      {/* Generate / Download */}
      {state.downloadId ? (
        <div className="text-center space-y-3">
          <div className="text-green-600 text-5xl">{"\u2705"}</div>
          <p className="font-semibold text-gray-800">Generation Complete!</p>
          <div className="flex items-center justify-center gap-3">
            <button
              onClick={handleDownload}
              className="inline-block bg-green-600 text-white px-6 py-3 rounded-lg font-medium hover:bg-green-700 transition"
            >
              Download Excel
            </button>
            {onViewDashboard && (
              <button
                onClick={onViewDashboard}
                className="inline-block bg-teal-600 text-white px-6 py-3 rounded-lg font-medium hover:bg-teal-700 transition"
              >
                View Dashboard
              </button>
            )}
          </div>
        </div>
      ) : (
        <button
          onClick={handleGenerate}
          disabled={state.isGenerating}
          className={`w-full py-3 rounded-lg font-semibold text-white transition
            ${state.isGenerating ? "bg-blue-400 cursor-wait" : "bg-blue-600 hover:bg-blue-700"}
          `}
        >
          {state.isGenerating ? (
            <span className="flex items-center justify-center gap-2">
              <svg
                className="animate-spin h-5 w-5 text-white"
                viewBox="0 0 24 24"
              >
                <circle
                  className="opacity-25"
                  cx="12"
                  cy="12"
                  r="10"
                  stroke="currentColor"
                  strokeWidth="4"
                  fill="none"
                />
                <path
                  className="opacity-75"
                  fill="currentColor"
                  d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"
                />
              </svg>
              Generating...
            </span>
          ) : (
            "Generate Data Pack"
          )}
        </button>
      )}
    </div>
  );
}

function Row({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex justify-between">
      <span className="text-gray-500">{label}</span>
      <span className="text-gray-800 font-medium text-right">{value}</span>
    </div>
  );
}
