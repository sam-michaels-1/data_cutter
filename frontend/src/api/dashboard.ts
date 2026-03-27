import axios from "axios";
import type { DashboardResponse } from "../types/dashboard";

const api = axios.create({ baseURL: "/api" });

export async function fetchDashboard(
  sessionId: string,
  granularity?: string
): Promise<DashboardResponse> {
  const params: Record<string, string> = {};
  if (granularity) params.granularity = granularity;
  const { data } = await api.get<DashboardResponse>(
    `/dashboard/${sessionId}`,
    { params }
  );
  return data;
}
