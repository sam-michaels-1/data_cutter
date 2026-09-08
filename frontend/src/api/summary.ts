/**
 * Summary data fetcher.
 * Computes the retention summary table from the loaded workbook.
 */
import type { SummaryResult } from "../engine/summary_compute";
import { getCurrentWorkbook, getCurrentConfig } from "./client";
import { computeSummaryData } from "../engine/summary_compute";

export async function fetchSummaryData(
  _sessionId: string,
  granularity?: string,
  filters?: Record<string, string | string[]>,
  identifier?: string,
): Promise<SummaryResult> {
  const wb = getCurrentWorkbook();
  const config = getCurrentConfig();

  if (!wb || !config) {
    throw new Error("No data loaded. Please import a file first.");
  }

  return computeSummaryData(wb, config, granularity, filters, identifier);
}
