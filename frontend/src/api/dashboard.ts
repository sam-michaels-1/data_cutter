import axios from "axios";
import type { DashboardResponse } from "../types/dashboard";

const api = axios.create({ baseURL: "/api" });

export async function fetchDashboard(
  sessionId: string,
  granularity?: string,
  filters?: Record<string, string>,
  topN?: number
): Promise<DashboardResponse> {
  const params: Record<string, string | number> = {};
  if (granularity) params.granularity = granularity;
  if (filters && Object.keys(filters).length > 0)
    params.filters = JSON.stringify(filters);
  if (topN != null) params.top_n = topN;
  const { data } = await api.get<DashboardResponse>(
    `/dashboard/${sessionId}`,
    { params }
  );
  return data;
}
