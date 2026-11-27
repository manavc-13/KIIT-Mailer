// /api/senders: the sender accounts the signed-in user may send from.
// Senders see only accounts allotted to them; admins see every enabled account.
import { Hono } from 'hono';
import { and, asc, eq } from 'drizzle-orm';
import { db, schema } from '../db/client.js';
import type { AppEnv } from '../types.js';
import { me, requireUser } from '../lib/session.js';
import { toSenderAccount, usageToday } from '../lib/senders.js';
import type { SenderAccount } from '../../shared/types.js';

const r = new Hono<AppEnv>();
r.use('*', requireUser);

r.get('/', async (c) => {
  const user = me(c);
  const rows = user.role === 'admin'
    ? await db.select().from(schema.senderAccounts)
      .where(eq(schema.senderAccounts.disabled, false))
      .orderBy(asc(schema.senderAccounts.displayName))
    : (await db.select({ s: schema.senderAccounts }).from(schema.senderAccounts)
      .innerJoin(schema.userSenderAccess, eq(schema.userSenderAccess.senderAccountId, schema.senderAccounts.id))
      .where(and(eq(schema.userSenderAccess.userId, user.id), eq(schema.senderAccounts.disabled, false)))
      .orderBy(asc(schema.senderAccounts.displayName))).map(x => x.s);
  const usage = await usageToday(rows.map(s => s.id));
  return c.json<{ senders: SenderAccount[] }>({ senders: rows.map(s => toSenderAccount(s, usage.get(s.id) ?? 0)) });
});

export default r;
