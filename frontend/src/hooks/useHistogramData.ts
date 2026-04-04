import { useState, useEffect, useCallback } from "react";
import type { HistogramResult } from "../engine/histograms";
import { fetchHistogramData } from "../api/histograms";

export interface HistogramRefetchOptions {
  filters?: Record<string, string | string[]>;
  mekkoXAxis?: string;
  mekkoYAxis?: string;
  gridXAxis?: string;
  gridYAxis?: string;
}

interface UseHistogramResult {
  data: HistogramResult | null;
  loading: boolean;
  error: string | null;
  refetch: (granularity?: string, opts?: HistogramRefetchOptions) => void;
}

export function useHistogramData(sessionId: string | null): UseHistogramResult {
  const [data, setData] = useState<HistogramResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const refetch = useCallback(
    async (granularity?: string, opts?: HistogramRefetchOptions) => {
      if (!sessionId) return;
      setLoading(true);
      setError(null);
      try {
        const result = await fetchHistogramData(
          sessionId,
          granularity,
          opts?.filters,
          opts?.mekkoXAxis,
          opts?.mekkoYAxis,
          opts?.gridXAxis,
          opts?.gridYAxis,
        );
        setData(result);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : "Failed to load histogram data";
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
