import { useState, useEffect, useCallback } from "react";
import type { SummaryResult } from "../engine/summary_compute";
import { fetchSummaryData } from "../api/summary";

export interface SummaryRefetchOptions {
  filters?: Record<string, string | string[]>;
  identifier?: string;
}

interface UseSummaryResult {
  data: SummaryResult | null;
  loading: boolean;
  error: string | null;
  refetch: (granularity?: string, opts?: SummaryRefetchOptions) => void;
}

export function useSummaryData(sessionId: string | null): UseSummaryResult {
  const [data, setData] = useState<SummaryResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const refetch = useCallback(
    async (granularity?: string, opts?: SummaryRefetchOptions) => {
      if (!sessionId) return;
      setLoading(true);
      setError(null);
      try {
        const result = await fetchSummaryData(
          sessionId,
          granularity,
          opts?.filters,
          opts?.identifier,
        );
        setData(result);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : "Failed to load summary data";
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
