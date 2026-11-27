// Authorization helpers for sender accounts.
import { and, eq } from 'drizzle-orm';
import { db, schema } from '../db/client.js';
import { forbidden, notFound } from './http.js';
import type { UserRow } from '../types.js';

export type SenderRow = typeof schema.senderAccounts.$inferSelect;

/**
 * Returns the sender account if `user` may send from it (admins: any enabled account;
 * senders: only accounts allotted via user_sender_access). Throws 403/404 otherwise.
 */
export async function assertSenderAccess(user: UserRow, senderAccountId: string): Promise<SenderRow> {
  const [acct] = await db.select().from(schema.senderAccounts).where(eq(schema.senderAccounts.id, senderAccountId)).limit(1);
  if (!acct) throw notFound('Sender account not found.');
  if (acct.disabled) throw forbidden('This sender account is disabled.');
  if (user.role === 'admin') return acct;
  const [grant] = await db.select().from(schema.userSenderAccess)
    .where(and(eq(schema.userSenderAccess.userId, user.id), eq(schema.userSenderAccess.senderAccountId, senderAccountId)))
    .limit(1);
  if (!grant) throw forbidden('You are not allowed to send from this account.');
  return acct;
}

/** Today's date in Asia/Kolkata as YYYY-MM-DD (quota day boundary). */
export function quotaDay(d = new Date()): string {
  return new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kolkata', year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
}
