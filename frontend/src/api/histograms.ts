/**
 * Histogram data fetcher.
 * Computes histogram/distribution data from the loaded workbook.
 */
import type { HistogramResult } from "../engine/histograms";
import { getCurrentWorkbook, getCurrentConfig } from "./client";
import { computeHistogramData } from "../engine/histograms";

export async function fetchHistogramData(
  _sessionId: string,
  granularity?: string,
  filters?: Record<string, string | string[]>,
  mekkoXAxis?: string,
  mekkoYAxis?: string,
  gridXAxis?: string,
  gridYAxis?: string,
): Promise<HistogramResult> {
  const wb = getCurrentWorkbook();
  const config = getCurrentConfig();

  if (!wb || !config) {
    throw new Error("No data loaded. Please import a file first.");
  }

  return computeHistogramData(wb, config, granularity, filters, mekkoXAxis, mekkoYAxis, gridXAxis, gridYAxis);
}
