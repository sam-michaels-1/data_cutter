import { useCallback } from "react";
import type { WizardState } from "../../types/wizard";
import { uploadFile } from "../../api/client";
import FileUpload from "../ui/FileUpload";

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

  const handleSheetSelect = useCallback(
    (sheet: string) => {
      dispatch({ type: "SET_SHEET", sheet });
    },
    [dispatch]
  );

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-semibold text-gray-800 mb-1">
          Upload Your Data
        </h2>
        <p className="text-gray-500 text-sm">
          Upload your Excel file with customer ARR or revenue data.
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
            Which sheet contains the data?
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
    </div>
  );
}
