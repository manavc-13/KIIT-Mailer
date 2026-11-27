// /api/admin/suppressions: addresses that every send job skips automatically
// (one-click unsubscribes, bounces, and addresses an administrator added by hand).
import { Hono } from 'hono';
import { desc, eq, ilike } from 'drizzle-orm';
import { z } from 'zod';
import { db, schema } from '../db/client.js';
import type { AppEnv } from '../types.js';
import { badRequest, body, notFound, query } from '../lib/http.js';
import { me } from '../lib/session.js';
import { audit } from '../lib/audit.js';
import { isValidEmail } from '../../shared/text.js';

const LIST_MAX = 5000;

const AddSchema = z.object({
  email: z.string().trim().toLowerCase().max(320).refine(isValidEmail, 'Enter a valid email address.'),
  note: z.string().trim().max(500, 'Keep the note under 500 characters.').optional().nullable(),
});
const ListSchema = z.object({ q: z.string().trim().max(320).optional() });

export interface SuppressionRow { email: string; reason: 'unsubscribed' | 'bounced' | 'manual'; note: string | null; at: string }

const r = new Hono<AppEnv>();

r.get('/', async (c) => {
  const { q } = query(c, ListSchema);
  const s = schema.suppressions;
  const rows = await db.select().from(s)
    .where(q ? ilike(s.email, `%${q.replace(/[%_\\]/g, m => '\\' + m)}%`) : undefined)
    .orderBy(desc(s.at))
    .limit(LIST_MAX);
  return c.json<{ suppressions: SuppressionRow[] }>({
    suppressions: rows.map(x => ({ email: x.email, reason: x.reason, note: x.note, at: x.at.toISOString() })),
  });
});

r.post('/', async (c) => {
  const actor = me(c);
  const { email, note } = await body(c, AddSchema);
  const inserted = await db.insert(schema.suppressions)
    .values({ email, reason: 'manual', note: note || null, createdBy: actor.id })
    .onConflictDoNothing()
    .returning({ email: schema.suppressions.email });
  if (inserted.length) {
    await audit(c, 'suppression.add', { targetType: 'suppression', targetId: email, meta: { email, note: note || null } });
  }
  return c.json({ ok: true as const, alreadySuppressed: inserted.length === 0 });
});

r.delete('/:email', async (c) => {
  // Hono has already percent-decoded the path parameter.
  const email = c.req.param('email').trim().toLowerCase();
  if (!email) throw badRequest('Invalid email address.');
  const deleted = await db.delete(schema.suppressions).where(eq(schema.suppressions.email, email))
    .returning({ email: schema.suppressions.email, reason: schema.suppressions.reason });
  if (!deleted.length) throw notFound('That address is not on the suppression list.');
  await audit(c, 'suppression.remove', { targetType: 'suppression', targetId: email, meta: { email, reason: deleted[0].reason } });
  return c.json({ ok: true as const });
});

export default r;
