// SMTP transport, message options, size math and error classification.
// Ported from legacy/lib/mailer.js. Pure apart from nodemailer + crypto (no DB),
// so it can be unit-tested.
//
// Size facts:
//   - Gmail / Google Workspace caps the total *encoded* message at 25 MB.
//     Base64 inflates attachment bytes by 4/3, so the usable raw attachment
//     budget is ~18 MB once headers/body are accounted for.
import { randomUUID } from 'node:crypto';
import nodemailer, { type Transporter } from 'nodemailer';
import type SMTPTransport from 'nodemailer/lib/smtp-transport/index.js';
import type Mail from 'nodemailer/lib/mailer/index.js';
import { open } from './crypto.js';
import { encodedSize } from '../../shared/text.js';
import type { SenderRow } from './access.js';

export const GMAIL_MAX_MESSAGE_BYTES = 25 * 1024 * 1024; // Gmail / Workspace hard cap (encoded)
export const SOFT_WARN_BYTES = 10 * 1024 * 1024; // above this many external servers refuse
export const MIME_OVERHEAD_BYTES = 96 * 1024; // headers, boilerplate, base64 line breaks
export const BASE64_RATIO = 4 / 3;
export const MAX_IMAGE_BYTES = 10 * 1024 * 1024;

/** Usable raw attachment budget once base64 inflation and MIME overhead are subtracted. */
export function maxRawAttachmentBytes(): number {
  return Math.floor((GMAIL_MAX_MESSAGE_BYTES - MIME_OVERHEAD_BYTES) / BASE64_RATIO);
}

/** Estimated encoded size of a message: body parts + base64 attachments. */
export function estimateMessageBytes(html: string, text: string, attachmentRawBytes: number): number {
  return Buffer.byteLength(html || '', 'utf8') + Buffer.byteLength(text || '', 'utf8') + encodedSize(attachmentRawBytes) + MIME_OVERHEAD_BYTES;
}

export type Transport = Transporter<SMTPTransport.SentMessageInfo>;

/** Pooled Gmail SMTP transport (one connection, reused for up to 100 messages). */
export function createTransport(user: string, pass: string): Transport {
  if (!user || !pass) throw new Error('Missing SMTP credentials.');
  return nodemailer.createTransport({
    host: 'smtp.gmail.com',
    port: 465,
    secure: true,
    pool: true,
    maxConnections: 1,
    maxMessages: 100,
    auth: { user, pass },
    // Keep every network step well inside the ~32 s claim budget of one /run call.
    connectionTimeout: 15_000,
    greetingTimeout: 10_000,
    socketTimeout: 25_000,
  } as SMTPTransport.Options);
}

/** Decrypt a sender account's App Password. AAD convention: the sender account id. */
export function senderCredentials(row: Pick<SenderRow, 'id' | 'email' | 'appPasswordEnc' | 'appPasswordIv' | 'appPasswordTag' | 'keyVersion'>): { user: string; pass: string } {
  let pass: string;
  try {
    pass = open({ enc: row.appPasswordEnc, iv: row.appPasswordIv, tag: row.appPasswordTag, keyVersion: row.keyVersion }, row.id);
  } catch {
    throw new Error(`The stored App Password for ${row.email} could not be decrypted. An administrator needs to re-enter it under Sender accounts.`);
  }
  return { user: row.email, pass };
}

export type ErrorKind = 'auth' | 'size' | 'quota' | 'throttle' | 'network' | 'recipient' | 'other';
export interface ClassifiedError { message: string; retryable: boolean; kind: ErrorKind }

/** Map raw SMTP / Nodemailer errors to plain English plus whether a retry is likely to help. */
export function classifyError(err: unknown): ClassifiedError {
  const e = (err ?? {}) as { response?: string; message?: string; responseCode?: number; code?: string };
  const raw = String(e.response || e.message || err || 'Unknown error');
  const code = e.responseCode ?? e.code;
  const is = (re: RegExp) => re.test(raw);

  if (code === 535 || code === 534 || code === 'EAUTH' || is(/\b53[45]\b|Username and Password not accepted|Application-specific password required|Invalid login/i)) {
    return {
      kind: 'auth', retryable: false,
      message: 'App Password rejected by Google. An administrator needs to generate a new App Password (Google Account → Security → App Passwords) and update this sender account.',
    };
  }
  if (is(/\b552\b|\b523\b|message too large|Message size exceeds/i)) {
    return { kind: 'size', retryable: false, message: "Message exceeds Gmail's 25 MB limit. Remove or shrink attachments." };
  }
  if (is(/550[- ]5\.4\.5|452[- ]?4\.2\.1|Daily user sending limit exceeded|Daily sending quota exceeded/i)) {
    return { kind: 'quota', retryable: false, message: "Google's daily sending limit for this account is used up. Resume tomorrow or use another account." };
  }
  if (is(/\b421\b|454[- ]?4\.7\.0|\b4\.7\.0\b|Temporary System Problem|too many login attempts|rate limit/i)) {
    return { kind: 'throttle', retryable: true, message: 'Gmail is throttling this connection. It will be retried automatically.' };
  }
  if (code === 'ECONNECTION' || code === 'ETIMEDOUT' || code === 'ESOCKET' || code === 'EDNS' || is(/ECONNECTION|ETIMEDOUT|ESOCKET|ENOTFOUND|ECONNRESET|ECONNREFUSED|EPIPE|Greeting never received|Connection closed|timed? ?out/i)) {
    return { kind: 'network', retryable: true, message: 'Network problem reaching smtp.gmail.com. It will be retried automatically.' };
  }
  if (code === 'EENVELOPE' || is(/\b55[03]\b.*(5\.1\.1|does not exist|no such user|user unknown|invalid recipient)|No recipients defined|Invalid recipient/i)) {
    return { kind: 'recipient', retryable: false, message: 'The recipient address was rejected (it may not exist). ' + raw.slice(0, 200) };
  }
  return { kind: 'other', retryable: false, message: raw.slice(0, 500) };
}

export interface MailAttachment {
  filename: string; content: Buffer; contentType?: string;
  /** For inline (CID) images: the Content-ID without angle brackets. */
  cid?: string;
  contentDisposition?: 'attachment' | 'inline';
}

export interface BuildMailInput {
  sender: Pick<SenderRow, 'email' | 'displayName' | 'replyTo'>;
  to: string; cc?: string | null; bcc?: string | null;
  subject: string; html: string; text: string;
  attachments?: MailAttachment[];
  /** One-click https unsubscribe URL (personal mode). */
  unsubscribeUrl?: string | null;
  /** Add List-Unsubscribe headers. With no URL, only the mailto form is sent (no one-click). */
  unsubscribe?: boolean;
  date?: Date;
}

/** Nodemailer message options. From is ALWAYS the sender account's own address. */
export function buildMailOptions(input: BuildMailInput): Mail.Options {
  const { sender } = input;
  const headers: Record<string, string> = {};
  if (input.unsubscribe) {
    const mailto = `<mailto:${sender.email}?subject=unsubscribe>`;
    const https = input.unsubscribeUrl && /^https?:\/\//i.test(input.unsubscribeUrl) ? input.unsubscribeUrl : null;
    if (https) {
      headers['List-Unsubscribe'] = `${mailto}, <${https}>`;
      headers['List-Unsubscribe-Post'] = 'List-Unsubscribe=One-Click';
    } else {
      headers['List-Unsubscribe'] = mailto;
    }
  }
  return {
    from: { name: sender.displayName || sender.email, address: sender.email },
    to: input.to,
    cc: input.cc || undefined,
    bcc: input.bcc || undefined,
    replyTo: sender.replyTo || undefined,
    subject: input.subject,
    html: input.html,
    text: input.text,
    messageId: '<' + randomUUID() + '@kiit.ac.in>',
    date: input.date ?? new Date(),
    headers,
    attachments: (input.attachments || []).map(a => ({
      filename: a.filename, content: a.content, contentType: a.contentType,
      ...(a.cid ? { cid: a.cid, contentDisposition: 'inline' as const } : {}),
    })),
    xMailer: false,
  } as Mail.Options;
}

/** Message-ID nodemailer actually sent (falls back to the one we set). */
export function sentMessageId(info: { messageId?: string } | null | undefined, opts: Mail.Options): string | null {
  return info?.messageId || (typeof opts.messageId === 'string' ? opts.messageId : null);
}
