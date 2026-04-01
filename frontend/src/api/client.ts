/**
 * API client — handles file upload, column detection, and workbook generation
 * using ExcelJS and the local engine (all in-browser).
 */
import ExcelJS from "exceljs";
import type {
  UploadResponse,
  DetectColumnsResponse,
  GenerateRequest,
  GenerateResponse,
} from "../types/wizard";
import { detectColumns as engineDetect, detectTableColumns as engineDetectTable } from "../engine/detect";
import type { DetectTableResult } from "../engine/detect";
import { buildEngineConfig } from "../engine/config_builder";
import { generateDataPack } from "../engine/generator";

// In-memory store for the current session's workbook
let currentWorkbook: ExcelJS.Workbook | null = null;
let currentConfig: import("../engine/types").EngineConfig | null = null;
let currentDownloadBlob: Blob | null = null;

export function getCurrentWorkbook(): ExcelJS.Workbook | null {
  return currentWorkbook;
}

export function getCurrentConfig(): import("../engine/types").EngineConfig | null {
  return currentConfig;
}

export async function uploadFile(file: File): Promise<UploadResponse> {
  const arrayBuffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);

  currentWorkbook = wb;

  const sheetNames = wb.worksheets.map((ws) => ws.name);
  const sessionId = `local-${Date.now()}`;

  return {
    session_id: sessionId,
    filename: file.name,
    sheet_names: sheetNames,
  };
}

export async function detectColumns(
  _sessionId: string,
  sheetName: string
): Promise<DetectColumnsResponse> {
  if (!currentWorkbook) throw new Error("No workbook loaded");

  const result = engineDetect(currentWorkbook, sheetName);

  return {
    columns: result.columns,
    detected_mapping: result.detected_mapping,
    row_count: result.row_count,
    auto_scale_factor: result.auto_scale_factor,
    detected_frequency: result.detected_frequency,
    header_row: result.header_row,
  };
}

export async function detectTableCols(
  _sessionId: string,
  sheetName: string
): Promise<DetectTableResult> {
  if (!currentWorkbook) throw new Error("No workbook loaded");
  return engineDetectTable(currentWorkbook, sheetName);
}

export async function generate(
  req: GenerateRequest
): Promise<GenerateResponse> {
  if (!currentWorkbook) throw new Error("No workbook loaded");

  const isCleaned = req.input_format === 'cleaned';

  // Build engine config
  const config = buildEngineConfig({
    wb: currentWorkbook,
    sheetName: req.sheet_name,
    dataType: req.data_type,
    dateCol: req.column_mapping.date_col,
    customerIdCol: req.column_mapping.customer_id_col,
    arrCol: req.column_mapping.arr_col,
    attributes: req.attributes,
    outputGranularities: req.output_granularities,
    fiscalYearEndMonth: req.fiscal_year_end_month,
    rowCount: currentWorkbook.getWorksheet(req.sheet_name)?.rowCount
      ? currentWorkbook.getWorksheet(req.sheet_name)!.rowCount - 1
      : 0,
    scaleFactor: 1, // Will be overridden below
    dataFrequency: req.data_frequency || undefined,
    inputFormat: req.input_format || 'raw',
    dateColumns: req.date_columns,
    headerRow: req.header_row,
    dateHeaderRow: req.date_header_row,
  });

  // Detect the actual row count by checking the data rows
  const ws = currentWorkbook.getWorksheet(req.sheet_name);
  const hRow = req.header_row || 1;
  if (ws) {
    let rowCount = 0;
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber <= hRow) return; // skip header and anything above
      // Check if the row has any data (don't assume column 1)
      let hasData = false;
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cell.value != null) hasData = true;
      });
      if (hasData) rowCount++;
    });
    config.raw_data_last_row = rowCount + hRow;
  }

  // Get scale factor from detection
  if (isCleaned) {
    const detectResult = engineDetectTable(currentWorkbook, req.sheet_name);
    config.scale_factor = detectResult.auto_scale_factor;
  } else {
    const detectResult = engineDetect(currentWorkbook, req.sheet_name);
    config.scale_factor = detectResult.auto_scale_factor;
  }

  currentConfig = config;

  // Generate the Excel workbook
  const outputWb = await generateDataPack(config, currentWorkbook);

  // Convert to blob
  const buffer = await outputWb.xlsx.writeBuffer();
  currentDownloadBlob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const downloadId = `download-${Date.now()}`;
  return { status: "success", download_id: downloadId };
}

export function getDownloadUrl(_downloadId: string): string {
  if (!currentDownloadBlob) return "#";
  return URL.createObjectURL(currentDownloadBlob);
}

export function getDownloadBlob(): Blob | null {
  return currentDownloadBlob;
}
