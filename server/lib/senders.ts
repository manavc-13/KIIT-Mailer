// Sender account helpers: DB row -> API shapes, app password handling and the
// SMTP login check. The app password is only ever decrypted server-side.
//
// Encryption convention (shared with the send engine): the app password is
// sealed with AAD = the sender account id, i.e.
//   seal(appPassword, row.id) / open({...}, row.id)
import { and, eq, inArray } from 'drizzle-orm';
import nodemailer from 'nodemailer';
import { db, schema } from '../db/client.js';
import { open } from './crypto.js';
import { iso } from './http.js';
import { quotaDay, type SenderRow } from './access.js';
import type { SenderAccount, SenderAccountAdmin, ThemeId } from '../../shared/types.js';

const THEME_IDS: ThemeId[] = ['heritage', 'meridian', 'letterhead', 'slate', 'plain'];
export const isThemeId = (v: unknown): v is ThemeId => typeof v === 'string' && (THEME_IDS as string[]).includes(v);

export function toSenderAccount(row: SenderRow, usedToday: number): SenderAccount {
  return {
    id: row.id,
    email: row.email,
    displayName: row.displayName,
    replyTo: row.replyTo,
    defaultTheme: isThemeId(row.defaultTheme) ? row.defaultTheme : 'heritage',
    signatureHtml: row.signatureHtml,
    signatureDoc: row.signatureDoc ?? null,
    dailyLimit: row.dailyLimit,
    usedToday,
    disabled: row.disabled,
  };
}

/** Admin view. Deliberately lists fields one by one so ciphertext columns can never leak. */
export function toSenderAccountAdmin(row: SenderRow, usedToday: number, userIds: string[]): SenderAccountAdmin {
  return {
    ...toSenderAccount(row, usedToday),
    appPasswordHint: row.appPasswordHint,
    lastVerifiedAt: iso(row.lastVerifiedAt),
    lastVerifyError: row.lastVerifyError,
    createdAt: row.createdAt.toISOString(),
    userIds,
  };
}

/** Recipients sent today (Asia/Kolkata day) per sender account id. */
export async function usageToday(ids: string[]): Promise<Map<string, number>> {
  const out = new Map<string, number>();
  if (!ids.length) return out;
  const rows = await db.select().from(schema.dailyUsage)
    .where(and(inArray(schema.dailyUsage.senderAccountId, ids), eq(schema.dailyUsage.day, quotaDay())));
  for (const r of rows) out.set(r.senderAccountId, r.recipientsSent);
  return out;
}

/** Decrypts the stored Google App Password. Server-side only; never put the result in a response. */
export function decryptAppPassword(row: SenderRow): string {
  return open({ enc: row.appPasswordEnc, iv: row.appPasswordIv, tag: row.appPasswordTag, keyVersion: row.keyVersion }, row.id);
}

/**
 * Google shows App Passwords as four groups of four letters ("abcd efgh ijkl mnop").
 * Strips the spaces and checks the shape. Returns the cleaned value or an error message.
 */
export function normalizeAppPassword(raw: string): { ok: true; value: string } | { ok: false; error: string } {
  const value = String(raw ?? '').replace(/\s+/g, '');
  if (!/^[a-zA-Z]{16}$/.test(value)) {
    return {
      ok: false,
      error: 'A Google App Password is 16 letters (Google shows it as four groups of four). Copy it again from myaccount.google.com/apppasswords.',
    };
  }
  return { ok: true, value };
}

const VERIFY_TIMEOUT_MS = 15_000;

/** Map an SMTP / network error from nodemailer to a plain-English explanation. */
export function explainSmtpError(err: unknown, email: string): string {
  const e = err as { code?: string | number; responseCode?: number; response?: string; message?: string } | null;
  const raw = String(e?.response || e?.message || err || '');
  const code = e?.responseCode ?? e?.code;
  const is = (re: RegExp) => re.test(raw);

  if (code === 'VERIFY_TIMEOUT') {
    return 'Gmail did not answer within 15 seconds. Check the network connection and try again.';
  }
  if (is(/Application-specific password required|5\.7\.9/i)) {
    return `Google requires an App Password for ${email}. Turn on 2-Step Verification for this account, then create an App Password at myaccount.google.com/apppasswords.`;
  }
  if (is(/web browser|5\.7\.14|WebLoginRequired/i)) {
    return `Google blocked the sign-in for ${email}. Sign in to Gmail in a browser once, confirm any security prompt, then try again.`;
  }
  if (code === 535 || code === 534 || code === 'EAUTH' || is(/535|534|Username and Password not accepted|BadCredentials|Invalid login/i)) {
    return `Google rejected the App Password for ${email}. Check that it was created for this exact account and copied correctly. App Passwords stop working if 2-Step Verification is turned off.`;
  }
  if (is(/454|4\.7\.0|too many login attempts|Temporary System Problem/i)) {
    return 'Gmail is temporarily refusing sign-ins for this account (too many attempts). Wait a few minutes and try again.';
  }
  if (code === 'ECONNECTION' || code === 'ETIMEDOUT' || code === 'ESOCKET' || code === 'EDNS'
    || is(/ECONNECTION|ETIMEDOUT|ESOCKET|ENOTFOUND|ECONNRESET|ECONNREFUSED|EAI_AGAIN|Greeting never received|Connection timeout/i)) {
    return 'Could not reach smtp.gmail.com. The server may be offline or blocking outgoing mail connections on port 465.';
  }
  return `Gmail refused the connection: ${raw.slice(0, 200) || 'unknown error'}`;
}

/** Signs in to smtp.gmail.com:465 with the given credentials (no mail is sent). */
export async function verifySmtpLogin(email: string, appPassword: string): Promise<{ ok: true } | { ok: false; error: string }> {
  const transporter = nodemailer.createTransport({
    host: 'smtp.gmail.com',
    port: 465,
    secure: true,
    auth: { user: email, pass: appPassword },
    connectionTimeout: VERIFY_TIMEOUT_MS,
    greetingTimeout: VERIFY_TIMEOUT_MS,
    socketTimeout: VERIFY_TIMEOUT_MS,
  });
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    await Promise.race([
      transporter.verify(),
      new Promise((_, reject) => {
        timer = setTimeout(() => reject(Object.assign(new Error('timeout'), { code: 'VERIFY_TIMEOUT' })), VERIFY_TIMEOUT_MS + 1000);
      }),
    ]);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: explainSmtpError(err, email) };
  } finally {
    if (timer) clearTimeout(timer);
    transporter.close();
  }
}
