/**
 * Generic data-fetching hook for Dataverse resources.
 * Provides loading / error / data state with a refresh callback.
 */

import { useCallback, useEffect, useRef, useState } from "react";

export interface UseDataverseResult<T> {
  data: T | null;
  isLoading: boolean;
  error: string | null;
  refresh: () => void;
}

export function useDataverse<T>(
  fetcher: () => Promise<T>,
  deps: unknown[] = []
): UseDataverseResult<T> {
  const [data, setData] = useState<T | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const mountedRef = useRef(true);

  const load = useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      const result = await fetcher();
      if (mountedRef.current) {
        setData(result);
      }
    } catch (err) {
      if (mountedRef.current) {
        setError(err instanceof Error ? err.message : String(err));
      }
    } finally {
      if (mountedRef.current) {
        setIsLoading(false);
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  useEffect(() => {
    mountedRef.current = true;
    load();
    return () => {
      mountedRef.current = false;
    };
  }, [load]);

  return { data, isLoading, error, refresh: load };
}
