// Append-only audit trail. Never throws: a failed audit write is logged, not surfaced.
import type { Context } from 'hono';
import { db, schema } from '../db/client.js';
import type { AuditAction } from '../../shared/types.js';
import { clientIp } from './http.js';
import type { AppEnv } from '../types.js';

export async function audit(
  c: Context<AppEnv> | null,
  action: AuditAction,
  opts: { targetType?: string; targetId?: string; meta?: Record<string, unknown>; actorId?: string | null; actorEmail?: string | null } = {},
): Promise<void> {
  try {
    const user = c?.get('user');
    await db.insert(schema.auditLog).values({
      actorId: opts.actorId !== undefined ? opts.actorId : user?.id ?? null,
      actorEmail: opts.actorEmail !== undefined ? opts.actorEmail : user?.email ?? null,
      action,
      targetType: opts.targetType ?? null,
      targetId: opts.targetId ?? null,
      meta: opts.meta ?? null,
      ip: c ? clientIp(c) : null,
    });
  } catch (err) {
    console.error('[audit] failed to write', action, err);
  }
}
