// /api/admin/senders: add, edit, rotate, verify and disable @kiit.ac.in sender accounts.
// App passwords are write-only: they are verified against Gmail, sealed with
// AES-256-GCM (AAD = account id) and never returned in any response.
import { Hono } from 'hono';
import crypto from 'node:crypto';
import { asc, eq, inArray } from 'drizzle-orm';
import { z } from 'zod';
import { db, schema } from '../db/client.js';
import type { AppEnv } from '../types.js';
import { badRequest, body, conflict, notFound } from '../lib/http.js';
import { me } from '../lib/session.js';
import { audit } from '../lib/audit.js';
import { seal } from '../lib/crypto.js';
import type { SenderRow } from '../lib/access.js';
import {
  decryptAppPassword, normalizeAppPassword, toSenderAccountAdmin, usageToday, verifySmtpLogin,
} from '../lib/senders.js';
import { isKiitEmail, isValidEmail } from '../../shared/text.js';
import type { SenderAccountAdmin } from '../../shared/types.js';

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const KiitEmail = z.string().trim().toLowerCase().max(320)
  .refine(isValidEmail, 'Enter a valid email address.')
  .refine(isKiitEmail, 'Sender accounts must be @kiit.ac.in addresses.');
const DisplayName = z.string().trim().min(1, 'Enter a display name.').max(120, 'Keep the display name under 120 characters.');
const ReplyTo = z.string().trim().toLowerCase().max(320)
  .refine(v => v === '' || isValidEmail(v), 'Enter a valid reply-to address, or leave it empty.')
  .transform(v => v || null)
  .nullable();
const DailyLimit = z.number().int('The daily limit must be a whole number.').min(1, 'The daily limit must be at least 1.')
  .max(2000, 'Google Workspace accounts can send to at most 2,000 recipients a day.');
const Theme = z.enum(['heritage', 'meridian', 'letterhead', 'slate', 'plain']);
const SignatureHtml = z.string().max(200_000, 'The signature is too large.').nullable();
const AppPassword = z.string().max(64);

const CreateSchema = z.object({
  email: KiitEmail, displayName: DisplayName, replyTo: ReplyTo.optional(), appPassword: AppPassword,
  dailyLimit: DailyLimit.optional(), defaultTheme: Theme.optional(),
  signatureHtml: SignatureHtml.optional(), signatureDoc: z.unknown().optional(),
});
const UpdateSchema = z.object({
  displayName: DisplayName.optional(), replyTo: ReplyTo.optional(), appPassword: AppPassword.optional(),
  dailyLimit: DailyLimit.optional(), defaultTheme: Theme.optional(),
  signatureHtml: SignatureHtml.optional(), signatureDoc: z.unknown().optional(), disabled: z.boolean().optional(),
});
const TestSchema = z.object({ email: KiitEmail, appPassword: AppPassword });

async function loadSender(id: string): Promise<SenderRow> {
  if (!UUID_RE.test(id)) throw notFound('Sender account not found.');
  const [row] = await db.select().from(schema.senderAccounts).where(eq(schema.senderAccounts.id, id)).limit(1);
  if (!row) throw notFound('Sender account not found.');
  return row;
}

async function adminViews(rows: SenderRow[]): Promise<SenderAccountAdmin[]> {
  const ids = rows.map(r => r.id);
  const usage = await usageToday(ids);
  const grants = new Map<string, string[]>();
  if (ids.length) {
    const g = await db.select().from(schema.userSenderAccess).where(inArray(schema.userSenderAccess.senderAccountId, ids));
    for (const x of g) grants.set(x.senderAccountId, [...(grants.get(x.senderAccountId) ?? []), x.userId]);
  }
  return rows.map(r => toSenderAccountAdmin(r, usage.get(r.id) ?? 0, grants.get(r.id) ?? []));
}

const adminView = async (row: SenderRow) => (await adminViews([row]))[0];

function cleanPassword(raw: string): string {
  const n = normalizeAppPassword(raw);
  if (!n.ok) throw badRequest(n.error, { field: 'appPassword' });
  return n.value;
}

const r = new Hono<AppEnv>();

r.get('/', async (c) => {
  const rows = await db.select().from(schema.senderAccounts).orderBy(asc(schema.senderAccounts.displayName));
  return c.json<{ senders: SenderAccountAdmin[] }>({ senders: await adminViews(rows) });
});

r.post('/test-credentials', async (c) => {
  const { email, appPassword } = await body(c, TestSchema);
  const n = normalizeAppPassword(appPassword);
  if (!n.ok) return c.json({ ok: false, error: n.error });
  const result = await verifySmtpLogin(email, n.value);
  return c.json(result.ok ? { ok: true } : { ok: false, error: result.error });
});

r.post('/', async (c) => {
  const actor = me(c);
  const input = await body(c, CreateSchema);
  const [dupe] = await db.select({ id: schema.senderAccounts.id }).from(schema.senderAccounts)
    .where(eq(schema.senderAccounts.email, input.email)).limit(1);
  if (dupe) throw conflict('This sender account has already been added.');

  const appPassword = cleanPassword(input.appPassword);
  const check = await verifySmtpLogin(input.email, appPassword);
  if (!check.ok) throw badRequest(check.error, { field: 'appPassword' });

  const id = crypto.randomUUID();
  const sealed = seal(appPassword, id);
  const now = new Date();
  let row: SenderRow;
  try {
    [row] = await db.insert(schema.senderAccounts).values({
      id,
      email: input.email,
      displayName: input.displayName,
      replyTo: input.replyTo ?? null,
      appPasswordEnc: sealed.enc,
      appPasswordIv: sealed.iv,
      appPasswordTag: sealed.tag,
      keyVersion: sealed.keyVersion,
      appPasswordHint: appPassword.slice(-4),
      dailyLimit: input.dailyLimit ?? 2000,
      defaultTheme: input.defaultTheme ?? 'heritage',
      signatureHtml: input.signatureHtml ?? null,
      signatureDoc: input.signatureDoc ?? null,
      lastVerifiedAt: now,
      lastVerifyError: null,
      createdBy: actor.id,
    }).returning();
  } catch (e) {
    const err = e as { code?: string; cause?: { code?: string } };
    if (err?.code === '23505' || err?.cause?.code === '23505') throw conflict('This sender account has already been added.');
    throw e;
  }
  await audit(c, 'sender.create', {
    targetType: 'sender_account', targetId: id,
    meta: { email: row.email, displayName: row.displayName, dailyLimit: row.dailyLimit, defaultTheme: row.defaultTheme },
  });
  return c.json<{ sender: SenderAccountAdmin }>({ sender: await adminView(row) }, 201);
});

r.patch('/:id', async (c) => {
  const row = await loadSender(c.req.param('id'));
  const input = await body(c, UpdateSchema);
  const set: Partial<typeof schema.senderAccounts.$inferInsert> = {};
  const changed: string[] = [];

  if (input.displayName !== undefined && input.displayName !== row.displayName) { set.displayName = input.displayName; changed.push('displayName'); }
  if (input.replyTo !== undefined && input.replyTo !== row.replyTo) { set.replyTo = input.replyTo; changed.push('replyTo'); }
  if (input.dailyLimit !== undefined && input.dailyLimit !== row.dailyLimit) { set.dailyLimit = input.dailyLimit; changed.push('dailyLimit'); }
  if (input.defaultTheme !== undefined && input.defaultTheme !== row.defaultTheme) { set.defaultTheme = input.defaultTheme; changed.push('defaultTheme'); }
  if (input.signatureHtml !== undefined) { set.signatureHtml = input.signatureHtml; changed.push('signature'); }
  if (input.signatureDoc !== undefined) { set.signatureDoc = input.signatureDoc ?? null; if (!changed.includes('signature')) changed.push('signature'); }

  let rotated = false;
  if (input.appPassword !== undefined && input.appPassword.trim() !== '') {
    const appPassword = cleanPassword(input.appPassword);
    const check = await verifySmtpLogin(row.email, appPassword);
    if (!check.ok) throw badRequest(check.error, { field: 'appPassword' });
    const sealed = seal(appPassword, row.id);
    Object.assign(set, {
      appPasswordEnc: sealed.enc, appPasswordIv: sealed.iv, appPasswordTag: sealed.tag, keyVersion: sealed.keyVersion,
      appPasswordHint: appPassword.slice(-4), lastVerifiedAt: new Date(), lastVerifyError: null,
    });
    rotated = true;
  }
  const disabledChanged = input.disabled !== undefined && input.disabled !== row.disabled;
  if (disabledChanged) set.disabled = input.disabled;

  if (Object.keys(set).length) {
    set.updatedAt = new Date();
    await db.update(schema.senderAccounts).set(set).where(eq(schema.senderAccounts.id, row.id));
  }
  const target = { targetType: 'sender_account', targetId: row.id };
  if (changed.length) {
    const values: Record<string, unknown> = {};
    for (const k of changed) if (k !== 'signature') values[k] = (set as Record<string, unknown>)[k];
    await audit(c, 'sender.update', { ...target, meta: { email: row.email, changed, values } });
  }
  if (rotated) await audit(c, 'sender.password_rotated', { ...target, meta: { email: row.email } });
  if (disabledChanged) await audit(c, input.disabled ? 'sender.disable' : 'sender.enable', { ...target, meta: { email: row.email } });

  return c.json<{ sender: SenderAccountAdmin }>({ sender: await adminView(await loadSender(row.id)) });
});

r.post('/:id/verify', async (c) => {
  const row = await loadSender(c.req.param('id'));
  let result: { ok: true } | { ok: false; error: string };
  try {
    result = await verifySmtpLogin(row.email, decryptAppPassword(row));
  } catch {
    result = { ok: false, error: 'The stored App Password could not be decrypted (the server encryption key may have changed). Enter the App Password again.' };
  }
  const now = new Date();
  await db.update(schema.senderAccounts)
    .set(result.ok ? { lastVerifiedAt: now, lastVerifyError: null } : { lastVerifyError: result.error })
    .where(eq(schema.senderAccounts.id, row.id));
  await audit(c, 'sender.verify', {
    targetType: 'sender_account', targetId: row.id,
    meta: { email: row.email, ok: result.ok, ...(result.ok ? {} : { error: result.error }) },
  });
  const sender = await adminView(await loadSender(row.id));
  return c.json<{ ok: boolean; error?: string; sender: SenderAccountAdmin }>(
    result.ok ? { ok: true, sender } : { ok: false, error: result.error, sender },
  );
});

export default r;
