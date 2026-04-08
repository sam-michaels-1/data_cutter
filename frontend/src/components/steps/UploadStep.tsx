import { useCallback, useState } from "react";
import type { WizardState } from "../../types/wizard";
import { uploadFile, detectTableCols } from "../../api/client";
import FileUpload from "../ui/FileUpload";

interface Props {
  state: WizardState;
  dispatch: React.Dispatch<any>;
}

export default function UploadStep({ state, dispatch }: Props) {
  const [loadingSample, setLoadingSample] = useState(false);

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

  const handleUseSampleData = useCallback(async () => {
    setLoadingSample(true);
    try {
      const res = await fetch("/sample-data.xlsx");
      const buf = await res.arrayBuffer();
      const file = new File([buf], "Sample Raw Data.xlsx", {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      // Upload to load workbook into memory
      const uploadRes = await uploadFile(file);
      const sheetName = uploadRes.sheet_names[0];

      // Run detection on the sheet
      const detect = await detectTableCols(uploadRes.session_id, sheetName);

      // Auto-select all detected attributes
      const selectedAttrs = detect.attribute_cols.map((a) => ({
        display_name: a.header,
        letter: a.letter,
      }));

      // Skip straight to the Generate step with everything pre-configured
      dispatch({
        type: "LOAD_SAMPLE",
        sessionId: uploadRes.session_id,
        filename: uploadRes.filename,
        sheetNames: uploadRes.sheet_names,
        selectedSheet: sheetName,
        columns: detect.columns,
        dateColumns: detect.date_columns,
        customerNameCol: detect.customer_id_col,
        attributes: detect.attribute_cols,
        scaleFactor: detect.auto_scale_factor,
        rowCount: detect.row_count,
        detectedFrequency: detect.detected_frequency as "monthly" | "quarterly",
        headerRow: detect.header_row,
        dateHeaderRow: detect.date_header_row,
        selectedAttributes: selectedAttrs,
      });
    } catch (err) {
      dispatch({
        type: "SET_ERROR",
        error: "Failed to load sample data",
      });
    } finally {
      setLoadingSample(false);
    }
  }, [dispatch]);

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
        isLoading={state.isLoading || loadingSample}
        filename={state.filename}
      />

      {!state.filename && (
        <div className="text-center">
          <p className="text-gray-400 text-xs mb-1">or</p>
          <button
            onClick={handleUseSampleData}
            disabled={state.isLoading || loadingSample}
            className="inline-flex items-center gap-2 px-5 py-2.5 bg-blue-800 hover:bg-blue-900 text-white text-sm font-medium rounded-lg shadow-sm transition disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {loadingSample ? "Loading sample data..." : "Try with sample data"}
          </button>
          <p className="text-gray-400 text-xs mt-2">
            500 customers, quarterly ARR with attributes
          </p>
        </div>
      )}

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
