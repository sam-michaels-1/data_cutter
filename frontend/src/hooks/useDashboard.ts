import { useState, useEffect, useCallback } from "react";
import type { DashboardResponse } from "../types/dashboard";
import { fetchDashboard } from "../api/dashboard";

interface UseDashboardResult {
  data: DashboardResponse | null;
  loading: boolean;
  error: string | null;
  refetch: (granularity?: string) => void;
}

export function useDashboard(sessionId: string | null): UseDashboardResult {
  const [data, setData] = useState<DashboardResponse | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const refetch = useCallback(
    async (granularity?: string) => {
      if (!sessionId) return;
      setLoading(true);
      setError(null);
      try {
        const result = await fetchDashboard(sessionId, granularity);
        setData(result);
      } catch (err: unknown) {
        const msg =
          err instanceof Error ? err.message : "Failed to load dashboard";
        setError(msg);
      } finally {
        setLoading(false);
      }
    },
    [sessionId]
  );

  useEffect(() => {
    if (sessionId) refetch();
  }, [sessionId, refetch]);

  return { data, loading, error, refetch };
}
