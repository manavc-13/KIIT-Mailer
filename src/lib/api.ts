// Tiny typed fetch wrapper for /api. Always sends the CSRF header and JSON;
// throws ApiClientError with the server's plain-English message.
import type { ApiErrorBody } from '../../shared/types';

export class ApiClientError extends Error {
  status: number; code?: string; details?: unknown;
  constructor(status: number, body: Partial<ApiErrorBody>) {
    super(body.error || `Request failed (HTTP ${status})`);
    this.status = status; this.code = body.code; this.details = body.details;
  }
}

type Json = Record<string, unknown> | unknown[] | object;

/** Shown for HTTP 413 (the platform rejects the request body before it reaches our server). */
export const TOO_LARGE_MESSAGE = 'This send is too large to submit in one go. Split the recipient list or shorten the message.';

let onUnauthorized: (() => void) | null = null;
/** Registered by AuthProvider: a 401 anywhere drops the user back to the login screen. */
export function setUnauthorizedHandler(fn: () => void) { onUnauthorized = fn; }

async function request<T>(method: string, path: string, body?: Json | BodyInit, init?: RequestInit): Promise<T> {
  const isRaw = body instanceof Blob || body instanceof ArrayBuffer || body instanceof FormData;
  const res = await fetch(path.startsWith('/api') ? path : `/api${path}`, {
    method,
    credentials: 'same-origin',
    ...init,
    headers: {
      'X-Requested-With': 'kiit-mailer',
      ...(body !== undefined && !isRaw ? { 'Content-Type': 'application/json' } : {}),
      ...(init?.headers || {}),
    },
    body: body === undefined ? undefined : isRaw ? (body as BodyInit) : JSON.stringify(body),
  });
  const ct = res.headers.get('content-type') || '';
  const data = ct.includes('application/json') ? await res.json().catch(() => ({})) : null;
  if (!res.ok) {
    if (res.status === 401 && !path.includes('/auth/login') && !path.includes('/auth/me')) onUnauthorized?.();
    if (res.status === 413) throw new ApiClientError(413, { error: TOO_LARGE_MESSAGE, code: data?.code ?? 'too_large' });
    throw new ApiClientError(res.status, data || { error: `Unexpected server response (HTTP ${res.status})` });
  }
  return (data ?? (await res.text())) as T;
}

export const api = {
  get: <T>(path: string, init?: RequestInit) => request<T>('GET', path, undefined, init),
  post: <T>(path: string, body?: Json | BodyInit, init?: RequestInit) => request<T>('POST', path, body ?? {}, init),
  put: <T>(path: string, body?: Json, init?: RequestInit) => request<T>('PUT', path, body ?? {}, init),
  patch: <T>(path: string, body?: Json, init?: RequestInit) => request<T>('PATCH', path, body ?? {}, init),
  del: <T>(path: string, init?: RequestInit) => request<T>('DELETE', path, undefined, init),
};

export function errorMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  return String(err);
}

/**
 * Toast title for a failed send / test send. The server tags quota and sign-in failures with a code
 * (`quota_exhausted`, `auth_failed`); its `error` text (errorMessage) is the plain-English detail.
 */
export function sendErrorTitle(err: unknown, fallback: string): string {
  if (err instanceof ApiClientError) {
    if (err.code === 'quota_exhausted') return 'Daily sending limit reached';
    if (err.code === 'auth_failed') return 'Gmail rejected the sender account’s App Password';
  }
  return fallback;
}
