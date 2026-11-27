// Debounced POST /api/render/preview with "keep the last good render" semantics.
import { useEffect, useRef, useState } from 'react';
import { keepPreviousData, useQuery } from '@tanstack/react-query';
import { api, ApiClientError } from '@/lib/api';
import type { RenderRequest, RenderResult } from '../../shared/types';

export function useDebouncedValue<T>(value: T, ms: number): T {
  const [debounced, setDebounced] = useState(value);
  useEffect(() => {
    if (Object.is(value, debounced)) return;
    const t = setTimeout(() => setDebounced(value), ms);
    return () => clearTimeout(t);
  }, [value, ms, debounced]);
  return debounced;
}

/** FNV-1a 32-bit hash, hex. Short, stable query keys for large request bodies. */
export function hashString(s: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16).padStart(8, '0') + s.length.toString(16);
}

/** Stable key for a render request; matches the requestKey passed to onRendered. */
export function renderRequestKey(request: RenderRequest | null): string | null {
  return request ? hashString(JSON.stringify(request)) : null;
}

export type PreviewProblem =
  | { kind: 'no-sender'; message: string }
  | { kind: 'error'; message: string };

function describeError(err: unknown): PreviewProblem {
  if (err instanceof ApiClientError) {
    const detail = JSON.stringify(err.details ?? '');
    const aboutSender = /sender/i.test(err.message) || /senderAccountId/.test(detail);
    if (aboutSender && (err.status === 400 || err.status === 404)) {
      return { kind: 'no-sender', message: 'Choose a sender to preview' };
    }
    if (err.status === 404 && !err.code) {
      return { kind: 'error', message: 'The preview service is not available yet.' };
    }
    return { kind: 'error', message: err.message };
  }
  if (err instanceof TypeError) return { kind: 'error', message: 'Could not reach the server to render the preview.' };
  return { kind: 'error', message: err instanceof Error ? err.message : String(err) };
}

export interface RenderPreviewState {
  /** The current render, or the last good one while updating / after an error. */
  result: RenderResult | null;
  /** True while a change is waiting for the debounce or a request is in flight. */
  updating: boolean;
  /** The shown result is older than the latest input (error or still updating). */
  stale: boolean;
  problem: PreviewProblem | null;
  initialLoading: boolean;
  refetch: () => void;
}

export function useRenderPreview(
  request: RenderRequest | null,
  opts: { debounceMs?: number; onRendered?: (r: RenderResult, requestKey: string) => void } = {},
): RenderPreviewState {
  const json = request ? JSON.stringify(request) : null;
  const debouncedJson = useDebouncedValue(json, opts.debounceMs ?? 400);
  const hasSender = !!request?.senderAccountId;
  const key = debouncedJson ? hashString(debouncedJson) : null;

  const query = useQuery({
    queryKey: ['render-preview', key],
    queryFn: ({ signal }) => api.post<RenderResult>('/render/preview', JSON.parse(debouncedJson!) as RenderRequest, { signal }),
    enabled: !!debouncedJson && hasSender,
    placeholderData: keepPreviousData,
    retry: false,
    staleTime: 60_000,
    gcTime: 5 * 60_000,
  });

  const [lastGood, setLastGood] = useState<RenderResult | null>(null);
  const onRenderedRef = useRef(opts.onRendered);
  onRenderedRef.current = opts.onRendered;

  useEffect(() => {
    if (query.data && !query.isPlaceholderData && key) {
      setLastGood(query.data);
      // `key` identifies the request this result was rendered from (see renderRequestKey).
      onRenderedRef.current?.(query.data, key);
    }
    // dataUpdatedAt changes on every successful fetch, including identical payloads.
  }, [query.data, query.isPlaceholderData, query.dataUpdatedAt]);

  const debouncing = json !== debouncedJson;
  let problem: PreviewProblem | null = null;
  if (!request || !hasSender) problem = { kind: 'no-sender', message: 'Choose a sender to preview' };
  else if (query.isError && !debouncing) problem = describeError(query.error);

  const result = query.data ?? lastGood;
  return {
    result,
    updating: hasSender && (debouncing || query.isFetching),
    stale: !!result && (query.isPlaceholderData || query.isError || debouncing),
    problem,
    initialLoading: hasSender && !result && !query.isError,
    refetch: () => { void query.refetch(); },
  };
}
