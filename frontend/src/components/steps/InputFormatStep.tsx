import { useCallback, useEffect } from "react";
import type { WizardState, InputFormat } from "../../types/wizard";
import { detectColumns, detectTableCols } from "../../api/client";
import ColumnMapper from "../ui/ColumnMapper";

interface Props {
  state: WizardState;
  dispatch: React.Dispatch<any>;
}

export default function InputFormatStep({ state, dispatch }: Props) {
  const handleFormatSelect = useCallback(
    async (format: InputFormat) => {
      dispatch({ type: "SET_INPUT_FORMAT", format });

      if (!state.sessionId || !state.selectedSheet) return;

      dispatch({ type: "SET_LOADING", loading: true });
      try {
        if (format === "raw") {
          const res = await detectColumns(state.sessionId, state.selectedSheet);
          dispatch({
            type: "DETECT_SUCCESS",
            columns: res.columns,
            mapping: res.detected_mapping,
            attributes: res.detected_mapping.attribute_cols,
            scaleFactor: res.auto_scale_factor,
            rowCount: res.row_count,
            detectedFrequency:
              res.detected_frequency === "monthly" ||
              res.detected_frequency === "quarterly"
                ? res.detected_frequency
                : null,
            headerRow: res.header_row,
          });
        } else {
          const res = await detectTableCols(
            state.sessionId,
            state.selectedSheet
          );
          dispatch({
            type: "DETECT_TABLE_SUCCESS",
            columns: res.columns,
            dateColumns: res.date_columns,
            customerNameCol: res.customer_id_col,
            attributes: res.attribute_cols,
            scaleFactor: res.auto_scale_factor,
            rowCount: res.row_count,
            detectedFrequency:
              res.detected_frequency === "monthly" ||
              res.detected_frequency === "quarterly"
                ? res.detected_frequency
                : null,
            headerRow: res.header_row,
            dateHeaderRow: res.date_header_row,
            fiscalLabeled: res.fiscal_labeled,
          });
        }
      } catch (err: any) {
        dispatch({
          type: "SET_ERROR",
          error: err instanceof Error ? err.message : "Detection failed",
        });
      }
    },
    [state.sessionId, state.selectedSheet, dispatch]
  );

  // Auto-detect columns when entering step 2 with a valid session but no detection data
  useEffect(() => {
    if (
      state.sessionId &&
      state.selectedSheet &&
      state.columns.length === 0 &&
      !state.isLoading
    ) {
      handleFormatSelect(state.inputFormat);
    }
  }, [state.sessionId, state.selectedSheet]);

  // Show all columns in the dropdown so the user can always override detection
  const nonDateColumns = state.columns;

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-semibold text-gray-800 mb-1">
          Data Format
        </h2>
        <p className="text-gray-500 text-sm">
          How is your data structured?
        </p>
      </div>

      <div className="grid grid-cols-2 gap-4">
        <button
          onClick={() => handleFormatSelect("raw")}
          disabled={state.isLoading}
          className={`p-4 rounded-xl border-2 text-left transition
            ${
              state.inputFormat === "raw"
                ? "border-teal-600 bg-teal-50"
                : "border-gray-200 hover:border-gray-300"
            }
          `}
        >
          <div className="font-semibold text-gray-800 mb-1">Raw Data</div>
          <p className="text-xs text-gray-500">
            One row per customer per time period (long format). Each row has a
            date, customer ID, and ARR/revenue value.
          </p>
        </button>

        <button
          onClick={() => handleFormatSelect("cleaned")}
          disabled={state.isLoading}
          className={`p-4 rounded-xl border-2 text-left transition
            ${
              state.inputFormat === "cleaned"
                ? "border-teal-600 bg-teal-50"
                : "border-gray-200 hover:border-gray-300"
            }
          `}
        >
          <div className="font-semibold text-gray-800 mb-1">Cleaned Table</div>
          <p className="text-xs text-gray-500">
            Customers as rows, time periods as columns (wide format). Each cell
            contains the ARR/revenue for that customer and period.
          </p>
        </button>
      </div>

      {state.isLoading && (
        <div className="text-center text-sm text-gray-500">
          Analyzing columns...
        </div>
      )}

      {/* Raw format: show column mapper */}
      {state.inputFormat === "raw" && state.confirmedMapping && state.columns.length > 0 && (
        <div>
          <h3 className="text-sm font-semibold text-gray-700 mb-2">
            Detected Column Mapping
          </h3>
          <ColumnMapper
            columns={state.columns}
            mapping={state.confirmedMapping}
            onChange={(m) =>
              dispatch({ type: "SET_CONFIRMED_MAPPING", mapping: m })
            }
          />
          <p className="text-xs text-gray-400 mt-2">
            {state.rowCount.toLocaleString()} data rows detected &middot; Scale
            factor: {state.scaleFactor.toLocaleString()}
          </p>
        </div>
      )}

      {/* Cleaned format: no period columns recognized — explain why Next is disabled */}
      {state.inputFormat === "cleaned" &&
        !state.isLoading &&
        state.columns.length > 0 &&
        state.dateColumns.length === 0 && (
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-4">
            <p className="text-sm font-semibold text-amber-800">
              No time-period columns recognized
            </p>
            <p className="text-xs text-amber-700 mt-1">
              We couldn&apos;t read any of your column headers as quarters or
              months, so there&apos;s nothing to analyze yet (that&apos;s why
              Next is disabled). Rename your period columns to a supported format
              and re-upload.
            </p>
            <p className="text-xs text-amber-700 mt-2">
              Supported examples:{" "}
              <span className="font-mono">Q1-FY26</span>,{" "}
              <span className="font-mono">Q3 2026</span>,{" "}
              <span className="font-mono">Q3-2026</span>,{" "}
              <span className="font-mono">1Q26</span>,{" "}
              <span className="font-mono">Jan-26</span>, or real dates like{" "}
              <span className="font-mono">3/31/2026</span>.
            </p>
          </div>
        )}

      {/* Cleaned format: show detected layout */}
      {state.inputFormat === "cleaned" && state.dateColumns.length > 0 && (
        <div className="space-y-4">
          <div className="bg-green-50 border border-green-200 rounded-lg p-3">
            <p className="text-sm text-green-800 font-medium">
              Detected {state.dateColumns.length} date columns
            </p>
            <p className="text-xs text-green-600 mt-1">
              {state.dateColumns.length > 6
                ? `${state.dateColumns.slice(0, 3).join(", ")} ... ${state.dateColumns.slice(-3).join(", ")}`
                : state.dateColumns.join(", ")}
            </p>
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Customer ID / Name Column
            </label>
            <select
              value={state.customerNameCol || ""}
              onChange={(e) =>
                dispatch({ type: "SET_CUSTOMER_NAME_COL", col: e.target.value })
              }
              className="w-full border border-gray-300 rounded-lg px-3 py-2 bg-white"
            >
              {nonDateColumns.map((col) => (
                <option key={col.letter} value={col.letter}>
                  {col.letter} — {col.header}
                  {state.dateColumns.includes(col.letter) ? " (date)" : ""}
                </option>
              ))}
            </select>
          </div>

          <p className="text-xs text-gray-400">
            {state.rowCount.toLocaleString()} data rows detected &middot; Scale
            factor: {state.scaleFactor.toLocaleString()}
          </p>
        </div>
      )}
    </div>
  );
}
