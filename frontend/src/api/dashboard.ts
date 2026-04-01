/**
 * Client-side dashboard computation.
 * Replaces the backend API call with local engine computation.
 */
import type { DashboardResponse } from "../types/dashboard";
import { getCurrentWorkbook, getCurrentConfig } from "./client";
import { computeDashboard } from "../engine/compute";

export async function fetchDashboard(
  _sessionId: string,
  granularity?: string,
  filters?: Record<string, string>,
  _topN?: number
): Promise<DashboardResponse> {
  const wb = getCurrentWorkbook();
  const config = getCurrentConfig();

  if (!wb || !config) {
    throw new Error("No data loaded. Please import a file first.");
  }

  const result = computeDashboard(wb, config, granularity, filters, _topN || 10);

  // Cast to DashboardResponse (the shape matches)
  return result as unknown as DashboardResponse;
}
