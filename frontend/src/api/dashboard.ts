/**
 * Dashboard data fetcher.
 * Computes dashboard metrics from the loaded workbook using the local engine.
 */
import type { DashboardResponse } from "../types/dashboard";
import { getCurrentWorkbook, getCurrentConfig } from "./client";
import { computeDashboard } from "../engine/compute";

export async function fetchDashboard(
  _sessionId: string,
  granularity?: string,
  filters?: Record<string, string | string[]>,
  _topN?: number
): Promise<DashboardResponse> {
  const wb = getCurrentWorkbook();
  const config = getCurrentConfig();

  if (!wb || !config) {
    throw new Error("No data loaded. Please import a file first.");
  }

  const result = computeDashboard(wb, config, granularity, filters, _topN || 10);

  return result as DashboardResponse;
}
