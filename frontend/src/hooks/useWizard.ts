import { useReducer, useCallback } from "react";
import type {
  WizardState,
  DataType,
  DataFrequency,
  Granularity,
  ColumnMapping,
  AttributeSelection,
  DetectedMapping,
  ColumnInfo,
  AttributeCol,
} from "../types/wizard";

const INITIAL_STATE: WizardState = {
  currentStep: 1,
  sessionId: null,
  filename: null,
  sheetNames: [],
  selectedSheet: null,
  columns: [],
  detectedMapping: null,
  confirmedMapping: null,
  detectedAttributes: [],
  scaleFactor: 1,
  rowCount: 0,
  detectedFrequency: null,
  dataFrequency: null,
  dataType: "arr",
  outputGranularities: [],
  fiscalYearEndMonth: 12,
  selectedAttributes: [],
  downloadId: null,
  isGenerating: false,
  isLoading: false,
  error: null,
};

type Action =
  | { type: "SET_STEP"; step: number }
  | {
      type: "UPLOAD_SUCCESS";
      sessionId: string;
      filename: string;
      sheetNames: string[];
    }
  | {
      type: "DETECT_SUCCESS";
      columns: ColumnInfo[];
      mapping: DetectedMapping;
      attributes: AttributeCol[];
      scaleFactor: number;
      rowCount: number;
      detectedFrequency: DataFrequency | null;
    }
  | { type: "SET_SHEET"; sheet: string }
  | { type: "SET_CONFIRMED_MAPPING"; mapping: ColumnMapping }
  | { type: "SET_DATA_FREQUENCY"; dataFrequency: DataFrequency }
  | { type: "SET_DATA_TYPE"; dataType: DataType }
  | { type: "TOGGLE_GRANULARITY"; granularity: Granularity }
  | { type: "SET_FISCAL_MONTH"; month: number }
  | { type: "SET_SELECTED_ATTRIBUTES"; attrs: AttributeSelection[] }
  | { type: "TOGGLE_ATTRIBUTE"; attr: AttributeCol }
  | { type: "RENAME_ATTRIBUTE"; letter: string; newName: string }
  | { type: "GENERATE_START" }
  | { type: "GENERATE_SUCCESS"; downloadId: string }
  | { type: "SET_LOADING"; loading: boolean }
  | { type: "SET_ERROR"; error: string | null }
  | { type: "RESET" };

function reducer(state: WizardState, action: Action): WizardState {
  switch (action.type) {
    case "SET_STEP":
      return { ...state, currentStep: action.step, error: null };

    case "UPLOAD_SUCCESS":
      return {
        ...state,
        sessionId: action.sessionId,
        filename: action.filename,
        sheetNames: action.sheetNames,
        isLoading: false,
        error: null,
      };

    case "SET_SHEET":
      return { ...state, selectedSheet: action.sheet };

    case "DETECT_SUCCESS": {
      const freq = action.detectedFrequency;
      return {
        ...state,
        columns: action.columns,
        detectedMapping: action.mapping,
        confirmedMapping: {
          date_col: action.mapping.date_col,
          customer_id_col: action.mapping.customer_id_col,
          arr_col: action.mapping.arr_col,
        },
        detectedAttributes: action.attributes,
        scaleFactor: action.scaleFactor,
        rowCount: action.rowCount,
        detectedFrequency: freq,
        dataFrequency: freq,
        isLoading: false,
      };
    }

    case "SET_CONFIRMED_MAPPING":
      return { ...state, confirmedMapping: action.mapping };

    case "SET_DATA_FREQUENCY":
      return { ...state, dataFrequency: action.dataFrequency };

    case "SET_DATA_TYPE":
      return { ...state, dataType: action.dataType };

    case "TOGGLE_GRANULARITY": {
      const grans = state.outputGranularities.includes(action.granularity)
        ? state.outputGranularities.filter((g) => g !== action.granularity)
        : [...state.outputGranularities, action.granularity];
      return { ...state, outputGranularities: grans };
    }

    case "SET_FISCAL_MONTH":
      return { ...state, fiscalYearEndMonth: action.month };

    case "SET_SELECTED_ATTRIBUTES":
      return { ...state, selectedAttributes: action.attrs };

    case "TOGGLE_ATTRIBUTE": {
      const exists = state.selectedAttributes.find(
        (a) => a.letter === action.attr.letter
      );
      const attrs = exists
        ? state.selectedAttributes.filter(
            (a) => a.letter !== action.attr.letter
          )
        : [
            ...state.selectedAttributes,
            {
              display_name: action.attr.header,
              letter: action.attr.letter,
            },
          ];
      return { ...state, selectedAttributes: attrs };
    }

    case "RENAME_ATTRIBUTE": {
      const attrs = state.selectedAttributes.map((a) =>
        a.letter === action.letter
          ? { ...a, display_name: action.newName }
          : a
      );
      return { ...state, selectedAttributes: attrs };
    }

    case "GENERATE_START":
      return { ...state, isGenerating: true, error: null };

    case "GENERATE_SUCCESS":
      return {
        ...state,
        isGenerating: false,
        downloadId: action.downloadId,
      };

    case "SET_LOADING":
      return { ...state, isLoading: action.loading };

    case "SET_ERROR":
      return {
        ...state,
        error: action.error,
        isLoading: false,
        isGenerating: false,
      };

    case "RESET":
      return INITIAL_STATE;

    default:
      return state;
  }
}

export function useWizard() {
  const [state, dispatch] = useReducer(reducer, INITIAL_STATE);

  const nextStep = useCallback(
    () => dispatch({ type: "SET_STEP", step: state.currentStep + 1 }),
    [state.currentStep]
  );
  const prevStep = useCallback(
    () => dispatch({ type: "SET_STEP", step: state.currentStep - 1 }),
    [state.currentStep]
  );
  const goToStep = useCallback(
    (step: number) => dispatch({ type: "SET_STEP", step }),
    []
  );

  return { state, dispatch, nextStep, prevStep, goToStep };
}
