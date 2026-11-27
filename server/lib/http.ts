// HTTP helpers shared by every route: typed errors, zod body/query parsing, client IP.
import type { Context } from 'hono';
import { HTTPException } from 'hono/http-exception';
import type { ContentfulStatusCode } from 'hono/utils/http-status';
import { z } from 'zod';

export class ApiError extends HTTPException {
  code?: string; details?: unknown;
  constructor(status: ContentfulStatusCode, message: string, code?: string, details?: unknown) {
    super(status, { message });
    this.code = code; this.details = details;
  }
}

export const badRequest = (msg: string, details?: unknown) => new ApiError(400, msg, 'bad_request', details);
export const unauthorized = (msg = 'Please sign in.') => new ApiError(401, msg, 'unauthorized');
export const forbidden = (msg = 'You do not have access to this.') => new ApiError(403, msg, 'forbidden');
export const notFound = (msg = 'Not found.') => new ApiError(404, msg, 'not_found');
export const conflict = (msg: string) => new ApiError(409, msg, 'conflict');
export const tooMany = (msg: string) => new ApiError(429, msg, 'rate_limited');
/** Daily recipient quota for a sender account would be exceeded. Clients branch on code 'quota_exhausted'. */
export const quotaExhausted = (msg: string, details?: unknown) => new ApiError(400, msg, 'quota_exhausted', details);
/** Google rejected the sender's App Password. Clients branch on code 'auth_failed'. */
export const authFailed = (msg: string, details?: unknown) => new ApiError(400, msg, 'auth_failed', details);

/** Parse and validate a JSON body. Throws 400 with zod issues on failure. */
export async function body<S extends z.ZodType>(c: Context, schema: S): Promise<z.infer<S>> {
  let raw: unknown;
  try { raw = await c.req.json(); } catch { throw badRequest('Request body must be JSON.'); }
  const r = schema.safeParse(raw);
  if (!r.success) throw badRequest(r.error.issues[0]?.message || 'Invalid request.', r.error.issues);
  return r.data;
}

export function query<S extends z.ZodType>(c: Context, schema: S): z.infer<S> {
  const r = schema.safeParse(c.req.query());
  if (!r.success) throw badRequest(r.error.issues[0]?.message || 'Invalid query.', r.error.issues);
  return r.data;
}

export function clientIp(c: Context): string | null {
  const xff = c.req.header('x-forwarded-for');
  if (xff) return xff.split(',')[0].trim();
  return c.req.header('x-real-ip') || null;
}

export const iso = (d: Date | null | undefined) => (d ? d.toISOString() : null);
