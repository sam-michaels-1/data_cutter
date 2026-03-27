// Types matching the backend Pydantic schemas

export interface UploadResponse {
  session_id: string;
  filename: string;
  sheet_names: string[];
}

export interface ColumnInfo {
  letter: string;
  header: string;
  sample_values: string[];
}

export interface AttributeCol {
  header: string;
  letter: string;
}

export interface DetectedMapping {
  date_col: string;
  customer_id_col: string;
  arr_col: string;
  attribute_cols: AttributeCol[];
}

export interface DetectColumnsResponse {
  columns: ColumnInfo[];
  detected_mapping: DetectedMapping;
  row_count: number;
  auto_scale_factor: number;
}

export interface ColumnMapping {
  date_col: string;
  customer_id_col: string;
  arr_col: string;
}

export interface AttributeSelection {
  display_name: string;
  letter: string;
}

export interface GenerateRequest {
  session_id: string;
  sheet_name: string;
  data_type: string;
  column_mapping: ColumnMapping;
  attributes: AttributeSelection[];
  output_granularities: string[];
  fiscal_year_end_month: number;
}

export interface GenerateResponse {
  status: string;
  download_id: string;
}

// Wizard state
export type DataType = "arr" | "revenue";
export type Granularity = "monthly" | "quarterly" | "annual";

export interface WizardState {
  currentStep: number;
  // Step 1
  sessionId: string | null;
  filename: string | null;
  sheetNames: string[];
  selectedSheet: string | null;
  columns: ColumnInfo[];
  detectedMapping: DetectedMapping | null;
  confirmedMapping: ColumnMapping | null;
  detectedAttributes: AttributeCol[];
  scaleFactor: number;
  rowCount: number;
  // Step 2
  dataType: DataType;
  // Step 3
  outputGranularities: Granularity[];
  fiscalYearEndMonth: number;
  // Step 4
  selectedAttributes: AttributeSelection[];
  // Step 5
  downloadId: string | null;
  isGenerating: boolean;
  // Shared
  isLoading: boolean;
  error: string | null;
}
