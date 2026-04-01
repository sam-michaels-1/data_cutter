/**
 * Shared types for the engine.
 */
import type { Worksheet } from 'exceljs';

export interface EngineConfig {
  raw_data_sheet: string;
  raw_data_first_row: number;
  raw_data_last_row: number;
  customer_id_col: string;
  date_col: string;
  arr_col: string;
  attributes: Record<string, string>;  // {display_name: raw_col_letter}
  time_granularity: string;
  output_granularities: string[];
  fiscal_year_end_month: number;
  scale_factor: number;
  filter_breakouts: FilterBreakout[];
  data_type: string;
  // Cleaned table fields
  input_format: 'raw' | 'cleaned';
  date_columns?: string[];            // column letters with date data (cleaned path)
  date_header_row?: number;           // row containing date headers (may differ from header row)
}

export interface FilterBreakout {
  title: string;
  filters: Record<string, string>;
}

export interface FilterBlock {
  title: string;
  filters: Record<string, string>;
}

export interface CleanTabResult {
  sheetName: string;
  layout: import('./utils').CleanLayout;
  firstDataRow: number;
  lastDataRow: number;
}

/** Workbook wrapper - ExcelJS Workbook typed more loosely for our usage */
export type WB = import('exceljs').Workbook;
export type WS = Worksheet;
