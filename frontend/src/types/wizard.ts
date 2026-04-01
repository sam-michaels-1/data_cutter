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
  detected_frequency: string | null;
  header_row: number;
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
  data_frequency: string | null;
  column_mapping: ColumnMapping;
  attributes: AttributeSelection[];
  output_granularities: string[];
  fiscal_year_end_month: number;
  // Cleaned table fields
  input_format?: InputFormat;
  date_columns?: string[];
  customer_name_col?: string | null;
  header_row?: number;
  date_header_row?: number;
}

export interface GenerateResponse {
  status: string;
  download_id: string;
}

// Wizard state
export type DataType = "arr" | "revenue";
export type DataFrequency = "monthly" | "quarterly";
export type Granularity = "monthly" | "quarterly" | "annual";
export type InputFormat = "raw" | "cleaned";

export interface WizardState {
  currentStep: number;
  // Step 1: Upload
  sessionId: string | null;
  filename: string | null;
  sheetNames: string[];
  selectedSheet: string | null;
  // Step 2: Input Format + Detection
  inputFormat: InputFormat;
  columns: ColumnInfo[];
  detectedMapping: DetectedMapping | null;
  confirmedMapping: ColumnMapping | null;
  detectedAttributes: AttributeCol[];
  scaleFactor: number;
  rowCount: number;
  detectedFrequency: DataFrequency | null;
  // Cleaned-table specific
  dateColumns: string[];
  customerNameCol: string | null;
  dateHeaderRow: number | null;  // row with date headers (if different from headerRow)
  // Shared: which row the headers were found in
  headerRow: number;
  // Step 3: Frequency
  dataFrequency: DataFrequency | null;
  // Step 4: Data Type
  dataType: DataType;
  // Step 5: Granularity
  outputGranularities: Granularity[];
  fiscalYearEndMonth: number;
  // Step 6: Identifiers
  selectedAttributes: AttributeSelection[];
  // Step 7: Review
  downloadId: string | null;
  isGenerating: boolean;
  // Shared
  isLoading: boolean;
  error: string | null;
}
