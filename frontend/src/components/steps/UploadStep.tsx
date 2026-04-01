import { useCallback } from "react";
import type { WizardState } from "../../types/wizard";
import { uploadFile, detectColumns } from "../../api/client";
import FileUpload from "../ui/FileUpload";
import ColumnMapper from "../ui/ColumnMapper";

interface Props {
  state: WizardState;
  dispatch: React.Dispatch<any>;
}

export default function UploadStep({ state, dispatch }: Props) {
  const handleFileSelect = useCallback(
    async (file: File) => {
      dispatch({ type: "SET_LOADING", loading: true });
      try {
        const res = await uploadFile(file);
        const sheets = res.sheet_names ?? [];
        dispatch({
          type: "UPLOAD_SUCCESS",
          sessionId: res.session_id,
          filename: res.filename,
          sheetNames: sheets,
        });
        // Auto-select if only one sheet
        if (sheets.length === 1) {
          dispatch({ type: "SET_SHEET", sheet: sheets[0] });
          await handleDetect(res.session_id, sheets[0]);
        }
      } catch (err: any) {
        dispatch({
          type: "SET_ERROR",
          error: err instanceof Error ? err.message : "Upload failed",
        });
      }
    },
    [dispatch]
  );

  const handleDetect = useCallback(
    async (sessionId: string, sheetName: string) => {
      dispatch({ type: "SET_LOADING", loading: true });
      try {
        const res = await detectColumns(sessionId, sheetName);
        dispatch({
          type: "DETECT_SUCCESS",
          columns: res.columns,
          mapping: res.detected_mapping,
          attributes: res.detected_mapping.attribute_cols,
          scaleFactor: res.auto_scale_factor,
          rowCount: res.row_count,
          detectedFrequency: (res.detected_frequency === "monthly" || res.detected_frequency === "quarterly")
            ? res.detected_frequency
            : null,
        });
      } catch (err: any) {
        dispatch({
          type: "SET_ERROR",
          error: err instanceof Error ? err.message : "Detection failed",
        });
      }
    },
    [dispatch]
  );

  const handleSheetSelect = useCallback(
    async (sheet: string) => {
      dispatch({ type: "SET_SHEET", sheet });
      if (state.sessionId) {
        await handleDetect(state.sessionId, sheet);
      }
    },
    [state.sessionId, dispatch, handleDetect]
  );

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-semibold text-gray-800 mb-1">
          Upload Your Data
        </h2>
        <p className="text-gray-500 text-sm">
          Upload your raw Excel file with customer ARR data.
        </p>
      </div>

      <FileUpload
        onFileSelect={handleFileSelect}
        isLoading={state.isLoading}
        filename={state.filename}
      />

      {/* Sheet picker */}
      {state.sheetNames?.length > 1 && (
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Which sheet contains the raw data?
          </label>
          <select
            value={state.selectedSheet || ""}
            onChange={(e) => handleSheetSelect(e.target.value)}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 bg-white"
          >
            <option value="" disabled>
              Select a sheet...
            </option>
            {state.sheetNames.map((s) => (
              <option key={s} value={s}>
                {s}
              </option>
            ))}
          </select>
        </div>
      )}

      {/* Column mapping */}
      {state.confirmedMapping && state.columns?.length > 0 && (
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
    </div>
  );
}
