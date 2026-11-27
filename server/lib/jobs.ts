// Server-side send engine: job creation, the paced runner, state changes and
// test sends. Pure planning rules live in ./send/plan.ts; rendering in ./render.
import type { Context } from 'hono';
import { and, desc, eq, inArray, lt, sql } from 'drizzle-orm';
import { db, schema } from '../db/client.js';
import type { AppEnv, UserRow } from '../types.js';
import type {
  ComposeContent, CreateJobRequest, Job, JobMessage, RunJobResponse, TestSendRequest, ThemeId,
} from '../../shared/types.js';
import { parseEmailList, placeholdersIn, uniqueEmails } from '../../shared/text.js';
import { assertSenderAccess, quotaDay, type SenderRow } from './access.js';
import { audit } from './audit.js';
import { ApiError, badRequest, forbidden, notFound, quotaExhausted, iso } from './http.js';
import {
  buildMailOptions, classifyError, createTransport, estimateMessageBytes, senderCredentials, sentMessageId,
  GMAIL_MAX_MESSAGE_BYTES, type MailAttachment, type Transport,
} from './mailer.js';
import { renderEmail } from './render/index.js';
import { inlineHostedImages, type HostedImage } from './send/images.js';
import {
  planPersonal, planBroadcast, totalRecipients, resolveUnsubscribe, checkQuota, firstAddress, paceDelayMs, chunk,
  type PlanResult,
} from './send/plan.js';
import { assertOwnedAttachments, loadHostedImages, resolveAttachments } from './attachments.js';
import { unsubscribeUrl } from './unsubscribe.js';

type JobRow = typeof schema.sendJobs.$inferSelect;
type MessageRow = typeof schema.sendMessages.$inferSelect;

export const MAX_MESSAGES_PER_JOB = 5000;
const MAX_ATTEMPTS = 3;

const sleep = (ms: number) => new Promise<void>(res => setTimeout(res, Math.max(0, ms)));
const fmt = (n: number) => n.toLocaleString('en-US');

// ---------------------------------------------------------------- run timing (pure)
// Production runs as one Vercel function with maxDuration 60 s (scripts/build-vercel.mjs).
// If Vercel kills a /run mid-send, the message is left 'sending' with an unknown outcome,
// so every /run must finish well inside that limit:
//   - new messages are only claimed during the first ~32 s of the request,
//   - each sendMail is abandoned after 18 s (retryable timeout),
//   - which leaves >= 7 s for the post-send DB writes (with retries) and the response,
//     plus slack for cold start before our clock starts.
export const FUNCTION_MAX_MS = 60_000;
export const SEND_TIMEOUT_MS = 18_000;
export const DEFAULT_CLAIM_BUDGET_MS = 32_000;
/** Upper bound for RUN_BUDGET_MS: at least 25 s must remain for the last send (18 s) and its writes. */
export const MAX_CLAIM_BUDGET_MS = 35_000;
/** How long to wait between checks when only another runner's in-flight messages remain. */
export const IN_FLIGHT_POLL_MS = 5_000;
/** Back-off between attempts to record a message Gmail already accepted (3 attempts in all). */
export const RECORD_SENT_BACKOFF_MS = [300, 1000];

/** Claim window (ms from the start of a /run). RUN_BUDGET_MS may shorten it, never lengthen it past 35 s. */
export function claimBudgetMs(raw: string | undefined = process.env.RUN_BUDGET_MS): number {
  const n = Number(raw);
  if (!raw || !Number.isFinite(n) || n <= 0) return DEFAULT_CLAIM_BUDGET_MS;
  return Math.min(n, MAX_CLAIM_BUDGET_MS);
}

/** How long to wait for another runner's in-flight messages before looking again (0 = stop now). */
export function inFlightWaitMs(now: number, claimUntil: number, poll = IN_FLIGHT_POLL_MS): number {
  return Math.max(0, Math.min(poll, claimUntil - now));
}

export class SendTimeoutError extends Error {
  constructor(ms: number) {
    super(`Gmail did not confirm this email within ${Math.round(ms / 1000)} seconds, so it timed out.`);
    this.name = 'SendTimeoutError';
  }
}

/** Resolve/reject with `p`, or reject with SendTimeoutError after `ms`. `p` keeps running (its outcome is ignored). */
export function withTimeout<T>(p: Promise<T>, ms: number): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<never>((_, reject) => { timer = setTimeout(() => reject(new SendTimeoutError(ms)), ms); });
  return Promise.race([p, timeout]).finally(() => clearTimeout(timer));
}

/** Run `fn`; on failure wait delays[i] and try again, up to delays.length + 1 attempts in all. */
export async function retryAsync<T>(
  fn: () => Promise<T>, delays: number[],
  opts: { wait?: (ms: number) => Promise<void>; onError?: (err: unknown, attempt: number) => void } = {},
): Promise<T> {
  const wait = opts.wait ?? sleep;
  for (let i = 0; ; i++) {
    try {
      return await fn();
    } catch (err) {
      opts.onError?.(err, i + 1);
      if (i >= delays.length) throw err;
      await wait(delays[i]);
    }
  }
}

// ---------------------------------------------------------------- claim / settle policy (pure)
/** A 'sending' row claimed longer ago than this belongs to a runner that died mid-send. */
export const STALE_CLAIM_MS = 3 * 60_000;
const STALE_CLAIM = sql.raw(`interval '${STALE_CLAIM_MS / 1000} seconds'`);
export const STALE_SENDING_ERROR = 'Delivery status unknown: the server stopped while this message was being sent. Check with the recipient before retrying.';
export const CANCELLED_ERROR = 'Cancelled before sending.';

/**
 * A stale claim may or may not have reached Gmail, so it is never resent automatically:
 * it becomes 'failed' (retryable, so the user can choose "Retry failed") with STALE_SENDING_ERROR.
 * Mirrors the SQL in sweepStaleClaims.
 */
export function isStaleClaim(claimedAt: Date | null, now: Date = new Date()): boolean {
  return !claimedAt || now.getTime() - claimedAt.getTime() > STALE_CLAIM_MS;
}

/**
 * What happens to a claimed message that was not sent:
 *  - skip:    the job was cancelled meanwhile and the failure was retryable or account-level
 *             (so it was not delivered): 'skipped' with CANCELLED_ERROR.
 *  - release: back to the queue without counting the attempt (account-level: auth, quota, job paused).
 *  - requeue: back to the queue for another attempt (retryable, attempts left).
 *  - fail:    final failure.
 */
export type SettleAction = 'skip' | 'release' | 'requeue' | 'fail';
export function settleAction(o: { retryable: boolean; accountLevel: boolean; attempts: number; jobStatus: string | null | undefined }): SettleAction {
  if (o.jobStatus === 'cancelled' && (o.retryable || o.accountLevel)) return 'skip';
  if (o.accountLevel) return 'release';
  if (o.retryable && o.attempts < MAX_ATTEMPTS) return 'requeue';
  return 'fail';
}

// ---------------------------------------------------------------- mapping
export function toJob(row: JobRow, senderEmail: string, userName?: string | null): Job {
  return {
    id: row.id, userId: row.userId, ...(userName != null ? { userName } : {}),
    senderAccountId: row.senderAccountId, senderEmail,
    templateId: row.templateId, mode: row.mode, subject: row.subject, status: row.status,
    totalMessages: row.totalMessages, totalRecipients: row.totalRecipients,
    sentCount: row.sentCount, failedCount: row.failedCount, skippedCount: row.skippedCount,
    createdAt: row.createdAt.toISOString(), startedAt: iso(row.startedAt), finishedAt: iso(row.finishedAt),
  };
}

export function toJobMessage(m: MessageRow): JobMessage {
  return {
    id: m.id, seq: m.seq, to: m.to, cc: m.cc, bcc: m.bcc, recipientCount: m.recipientCount,
    status: m.status, attempts: m.attempts, error: m.error, retryable: m.retryable, sentAt: iso(m.sentAt),
  };
}

async function loadJobRow(id: string): Promise<{ job: JobRow; senderEmail: string; userName: string } | null> {
  const [r] = await db.select({ job: schema.sendJobs, senderEmail: schema.senderAccounts.email, userName: schema.users.name })
    .from(schema.sendJobs)
    .innerJoin(schema.senderAccounts, eq(schema.senderAccounts.id, schema.sendJobs.senderAccountId))
    .innerJoin(schema.users, eq(schema.users.id, schema.sendJobs.userId))
    .where(eq(schema.sendJobs.id, id)).limit(1);
  return r ?? null;
}

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Load a job the user may see/act on (owner or admin). 404 otherwise. */
export async function getJobForUser(user: UserRow, id: string) {
  if (!UUID_RE.test(id)) throw notFound('Send job not found.');
  const r = await loadJobRow(id);
  if (!r || (r.job.userId !== user.id && user.role !== 'admin')) throw notFound('Send job not found.');
  return r;
}

async function jobView(id: string): Promise<Job> {
  const r = await loadJobRow(id);
  if (!r) throw notFound('Send job not found.');
  return toJob(r.job, r.senderEmail, r.userName);
}

export async function usedToday(senderAccountId: string): Promise<number> {
  const [u] = await db.select({ n: schema.dailyUsage.recipientsSent }).from(schema.dailyUsage)
    .where(and(eq(schema.dailyUsage.senderAccountId, senderAccountId), eq(schema.dailyUsage.day, quotaDay()))).limit(1);
  return u?.n ?? 0;
}

/**
 * Atomically reserve `n` recipients of today's quota (Asia/Kolkata day) BEFORE sending.
 * Returns the quota day reserved against (pass it to refundQuota), or null if the
 * reservation would exceed `dailyLimit`. Concurrent runners cannot overshoot the limit.
 */
export async function reserveQuota(senderAccountId: string, n: number, dailyLimit: number): Promise<string | null> {
  const day = quotaDay();
  if (n <= 0) return day;
  const u = schema.dailyUsage;
  await db.insert(u).values({ senderAccountId, day, recipientsSent: 0 }).onConflictDoNothing();
  const [row] = await db.update(u)
    .set({ recipientsSent: sql`${u.recipientsSent} + ${n}` })
    .where(and(eq(u.senderAccountId, senderAccountId), eq(u.day, day), sql`${u.recipientsSent} + ${n} <= ${dailyLimit}`))
    .returning({ n: u.recipientsSent });
  return row ? day : null;
}

/** Give back a reservation that was not used (send failed or never attempted). Never throws. */
export async function refundQuota(senderAccountId: string, day: string, n: number): Promise<void> {
  if (n <= 0) return;
  const u = schema.dailyUsage;
  try {
    await db.update(u).set({ recipientsSent: sql`greatest(${u.recipientsSent} - ${n}, 0)` })
      .where(and(eq(u.senderAccountId, senderAccountId), eq(u.day, day)));
  } catch (err) {
    // Over-counting is the safe direction: the sender just gets slightly less quota today.
    console.error(`[jobs] quota refund of ${n} for sender ${senderAccountId} on ${day} failed:`, (err as Error).message);
  }
}

/** Lower-cased suppressed addresses among `emails` (queried in chunks). */
export async function suppressedAmong(emails: string[]): Promise<Set<string>> {
  const all = [...new Set(emails.map(e => e.trim().toLowerCase()).filter(Boolean))];
  const out = new Set<string>();
  for (const part of chunk(all, 1000)) {
    const rows = await db.select({ e: schema.suppressions.email }).from(schema.suppressions).where(inArray(schema.suppressions.email, part));
    rows.forEach(r => out.add(r.e.toLowerCase()));
  }
  return out;
}

const listPreview = (list: string[], n = 5) => list.slice(0, n).join(', ') + (list.length > n ? ` and ${fmt(list.length - n)} more` : '');

// ---------------------------------------------------------------- create
export async function createJob(c: Context<AppEnv>, user: UserRow, req: CreateJobRequest): Promise<{ job: Job; warnings: string[] }> {
  const sender = await assertSenderAccess(user, req.senderAccountId);
  const warnings: string[] = [];

  const candidates = req.recipients.flatMap(r => [r.to, r.cc, r.bcc]).concat([req.globalCc ?? '', req.globalBcc ?? '']).flatMap(v => parseEmailList(v));
  const suppressed = await suppressedAmong(candidates);

  let plan: PlanResult;
  if (req.mode === 'bcc_broadcast') {
    plan = planBroadcast(req.recipients, sender.email, { globalCc: req.globalCc, globalBcc: req.globalBcc, suppressed });
  } else {
    plan = planPersonal(req.recipients, { globalCc: req.globalCc, globalBcc: req.globalBcc, dedupe: req.dedupe, suppressed });
  }

  if (plan.invalid.length) {
    throw badRequest(`${fmt(plan.invalid.length)} email address${plan.invalid.length > 1 ? 'es are' : ' is'} not valid: ${listPreview(plan.invalid)}. Fix ${plan.invalid.length > 1 ? 'them' : 'it'} and try again.`, { invalid: plan.invalid.slice(0, 50) });
  }
  if (plan.emptyRows.length) {
    throw badRequest(`Recipient row${plan.emptyRows.length > 1 ? 's' : ''} ${listPreview(plan.emptyRows.map(String), 8)} ${plan.emptyRows.length > 1 ? 'have' : 'has'} no email address.`);
  }
  const sendable = plan.messages.filter(m => !m.skipReason);
  if (!sendable.length) {
    throw badRequest(plan.suppressedAddresses.length
      ? 'Every recipient on this list has unsubscribed or is on the suppression list. Nothing would be sent.'
      : 'There are no recipients to send to.');
  }
  if (plan.messages.length > MAX_MESSAGES_PER_JOB) {
    throw badRequest(`This send would create ${fmt(plan.messages.length)} emails; the limit is ${fmt(MAX_MESSAGES_PER_JOB)} per send. Split the list into smaller sends.`);
  }

  const needed = totalRecipients(plan.messages);
  const quota = checkQuota(needed, sender.dailyLimit, await usedToday(sender.id), sender.email);
  if (!quota.ok) throw quotaExhausted(quota.message!, { needed, remaining: quota.remaining });

  const attachments = await assertOwnedAttachments(req.attachmentIds || [], user.id);

  let templateId: string | null = null;
  if (req.templateId && UUID_RE.test(req.templateId)) {
    const [t] = await db.select({ id: schema.templates.id }).from(schema.templates).where(eq(schema.templates.id, req.templateId)).limit(1);
    templateId = t?.id ?? null;
  }

  const unsubscribeEnabled = resolveUnsubscribe(req.unsubscribe, plan.messages);
  if (req.unsubscribe === 'auto' && unsubscribeEnabled) {
    warnings.push('More than 50 recipients are outside kiit.ac.in, so an unsubscribe link and one-click unsubscribe header were added.');
  }
  if (plan.suppressedAddresses.length) {
    warnings.push(`${fmt(plan.suppressedAddresses.length)} address${plan.suppressedAddresses.length > 1 ? 'es were' : ' was'} left out because ${plan.suppressedAddresses.length > 1 ? 'they' : 'it'} unsubscribed or ${plan.suppressedAddresses.length > 1 ? 'are' : 'is'} on the suppression list.`);
  }
  if (plan.duplicatesRemoved) warnings.push(`${fmt(plan.duplicatesRemoved)} duplicate recipient${plan.duplicatesRemoved > 1 ? 's were' : ' was'} removed.`);
  if (req.mode === 'bcc_broadcast') {
    const ph = placeholdersIn(req.subject, req.preheader, req.bodyHtml, req.bodyText);
    if (ph.length) warnings.push(`BCC broadcast sends one identical email to everyone, so ${ph.slice(0, 4).map(p => `{${p}}`).join(', ')} cannot be filled in per recipient.`);
    if (plan.messages.length > 1) warnings.push(`The list was split into ${fmt(plan.messages.length)} emails of at most 90 BCC recipients each (Gmail's per-message limit is 100).`);
  }

  const skipped = plan.messages.filter(m => m.skipReason).length;
  const jobId = await db.transaction(async (tx) => {
    const [job] = await tx.insert(schema.sendJobs).values({
      userId: user.id, senderAccountId: sender.id, templateId, mode: req.mode,
      subject: req.subject, preheader: req.preheader || '', editorMode: req.editorMode,
      bodyHtml: req.bodyHtml, bodyText: req.bodyText || null, theme: req.theme, includeSignature: req.includeSignature,
      globalCc: req.globalCc || null, globalBcc: req.globalBcc || null, unsubscribeEnabled,
      status: 'queued', totalMessages: plan.messages.length, totalRecipients: needed, skippedCount: skipped,
    }).returning({ id: schema.sendJobs.id });
    const rows = plan.messages.map((m, i) => ({
      jobId: job.id, seq: i + 1, to: m.to, cc: m.cc, bcc: m.bcc, recipientCount: m.recipientCount, rowData: m.rowData,
      status: (m.skipReason ? 'skipped' : 'queued') as 'skipped' | 'queued', error: m.skipReason ?? null,
    }));
    for (const part of chunk(rows, 500)) await tx.insert(schema.sendMessages).values(part);
    if (attachments.length) await tx.insert(schema.jobAttachments).values(attachments.map(a => ({ jobId: job.id, uploadId: a.id })));
    return job.id;
  });

  await audit(c, 'job.create', {
    targetType: 'job', targetId: jobId,
    meta: { senderAccountId: sender.id, senderEmail: sender.email, mode: req.mode, messages: plan.messages.length, recipients: needed, skipped, unsubscribe: unsubscribeEnabled, attachments: attachments.length },
  });
  return { job: await jobView(jobId), warnings };
}

// ---------------------------------------------------------------- run
function contentOf(job: JobRow): ComposeContent {
  return {
    subject: job.subject, preheader: job.preheader, editorMode: job.editorMode, bodyHtml: job.bodyHtml,
    bodyText: job.bodyText, theme: job.theme as ThemeId, includeSignature: job.includeSignature,
  };
}

/** Claim the next queued message. 'sending' rows are never reclaimed (see sweepStaleClaims). */
async function claimNext(jobId: string): Promise<MessageRow | null> {
  const m = schema.sendMessages;
  const [row] = await db.update(m)
    .set({ status: 'sending', claimedAt: sql`now()`, attempts: sql`${m.attempts} + 1` })
    .where(and(
      eq(m.id, sql`(SELECT id FROM send_messages WHERE job_id = ${jobId} AND status = 'queued' ORDER BY seq LIMIT 1 FOR UPDATE SKIP LOCKED)`),
      eq(m.status, 'queued'),
    ))
    .returning();
  return row ?? null;
}

/** Messages of this job currently claimed by a runner (after the sweep, all of them are fresh). */
async function inFlightCount(jobId: string): Promise<number> {
  const [r] = await db.select({ n: sql<number>`count(*)::int` }).from(schema.sendMessages)
    .where(and(eq(schema.sendMessages.jobId, jobId), eq(schema.sendMessages.status, 'sending')));
  return r?.n ?? 0;
}

/**
 * 'sending' rows claimed more than STALE_CLAIM_MS ago belong to a runner that died mid-send
 * (function timeout, crash). The email may or may not have been delivered, so they are NEVER
 * resent automatically: they become 'failed' (retryable, so "Retry failed" can pick them up
 * by the user's choice) with STALE_SENDING_ERROR, and count as failed.
 */
async function sweepStaleClaims(jobId: string): Promise<number> {
  const m = schema.sendMessages;
  const n = await db.transaction(async (tx) => {
    const rows = await tx.update(m)
      .set({ status: 'failed', retryable: true, error: STALE_SENDING_ERROR })
      .where(and(eq(m.jobId, jobId), eq(m.status, 'sending'), sql`(${m.claimedAt} IS NULL OR ${m.claimedAt} < now() - ${STALE_CLAIM})`))
      .returning({ id: m.id });
    if (rows.length) {
      await tx.update(schema.sendJobs).set({ failedCount: sql`${schema.sendJobs.failedCount} + ${rows.length}` }).where(eq(schema.sendJobs.id, jobId));
    }
    return rows.length;
  });
  if (n) console.warn(`[jobs] job ${jobId}: ${n} stale 'sending' message(s) marked failed (delivery status unknown).`);
  return n;
}

/**
 * Record a message Gmail accepted. Guarded by status = 'sending', so a retry after a commit
 * whose acknowledgement was lost does not double-count. Quota was reserved before sending.
 */
async function recordSent(jobId: string, msg: MessageRow, messageId: string | null): Promise<MessageRow> {
  const m = schema.sendMessages;
  return db.transaction(async (tx) => {
    const [r] = await tx.update(m)
      .set({ status: 'sent', smtpMessageId: messageId, sentAt: new Date(), error: null, retryable: null })
      .where(and(eq(m.id, msg.id), eq(m.status, 'sending'))).returning();
    if (r) {
      await tx.update(schema.sendJobs).set({ sentCount: sql`${schema.sendJobs.sentCount} + 1` }).where(eq(schema.sendJobs.id, jobId));
      return r;
    }
    const [cur] = await tx.select().from(m).where(eq(m.id, msg.id)).limit(1);
    return cur ?? msg;
  });
}

/**
 * Settle a claimed message that was NOT sent (see settleAction). The job row is locked first,
 * so the decision is atomic with cancel() and the counters always add up. Every write is
 * guarded by status = 'sending'.
 */
async function settleUnsent(
  jobId: string, msg: MessageRow, f: { message: string | null; retryable: boolean; accountLevel: boolean },
): Promise<{ row: MessageRow; action: SettleAction }> {
  const m = schema.sendMessages; const j = schema.sendJobs;
  return db.transaction(async (tx) => {
    const [cur] = await tx.select({ status: j.status }).from(j).where(eq(j.id, jobId)).for('update');
    const action = settleAction({ retryable: f.retryable, accountLevel: f.accountLevel, attempts: msg.attempts, jobStatus: cur?.status });
    const guard = and(eq(m.id, msg.id), eq(m.status, 'sending'));
    let rows: MessageRow[];
    if (action === 'skip') {
      rows = await tx.update(m).set({ status: 'skipped', error: CANCELLED_ERROR, retryable: null, claimedAt: null }).where(guard).returning();
      if (rows.length) await tx.update(j).set({ skippedCount: sql`${j.skippedCount} + 1` }).where(eq(j.id, jobId));
    } else if (action === 'release') {
      rows = await tx.update(m).set({ status: 'queued', attempts: Math.max(0, msg.attempts - 1), claimedAt: null, error: f.message }).where(guard).returning();
    } else if (action === 'requeue') {
      rows = await tx.update(m).set({ status: 'queued', claimedAt: null, error: f.message, retryable: true }).where(guard).returning();
    } else {
      rows = await tx.update(m).set({ status: 'failed', error: f.message, retryable: f.retryable }).where(guard).returning();
      if (rows.length) await tx.update(j).set({ failedCount: sql`${j.failedCount} + 1` }).where(eq(j.id, jobId));
    }
    const row = rows[0] ?? (await tx.select().from(m).where(eq(m.id, msg.id)).limit(1))[0] ?? msg;
    return { row, action };
  });
}

async function pauseJob(jobId: string): Promise<void> {
  await db.update(schema.sendJobs).set({ status: 'paused' })
    .where(and(eq(schema.sendJobs.id, jobId), inArray(schema.sendJobs.status, ['queued', 'sending'])));
}

async function jobStatus(jobId: string): Promise<string | null> {
  const [cur] = await db.select({ status: schema.sendJobs.status }).from(schema.sendJobs).where(eq(schema.sendJobs.id, jobId)).limit(1);
  return cur?.status ?? null;
}

/** Prepares per-message HTML: hosted images become CID attachments (cached across the run). */
class ImageCache {
  private cache = new Map<string, HostedImage | null>();
  async inline(html: string, ids: string[]): Promise<{ html: string; attachments: MailAttachment[] }> {
    const missing = ids.filter(id => !this.cache.has(id));
    if (missing.length) {
      const loaded = await loadHostedImages(missing);
      for (const id of missing) this.cache.set(id, loaded.find(l => l.id.toLowerCase() === id) ?? null);
    }
    const imgs = ids.map(id => this.cache.get(id)).filter((x): x is HostedImage => !!x);
    return imgs.length ? inlineHostedImages(html, imgs) : { html, attachments: [] };
  }
}

export type RunResult = RunJobResponse & { error?: string };

type SendOutcome = { ok: true; messageId: string | null } | { ok: false; message: string; retryable: boolean; kind: string };

export async function runJob(c: Context<AppEnv>, user: UserRow, jobId: string): Promise<RunResult> {
  const started = Date.now();
  const claimUntil = started + claimBudgetMs();
  const { job } = await getJobForUser(user, jobId);

  // Messages whose runner died mid-send become "Delivery status unknown" failures (never resent).
  await sweepStaleClaims(jobId);

  if (job.status === 'completed' || job.status === 'cancelled') {
    return { job: await jobView(jobId), processed: [], done: true };
  }

  const processed: MessageRow[] = [];
  const stop = async (error: string, extra: Partial<RunResult> = {}): Promise<RunResult> => {
    await pauseJob(jobId);
    return { job: await jobView(jobId), processed: processed.map(toJobMessage), done: true, error, ...extra };
  };

  // The sender may have been disabled, or the user's access revoked, since the job started:
  // pause with a plain-English reason rather than throwing 403 and leaving the job 'sending'.
  const [sender] = await db.select().from(schema.senderAccounts).where(eq(schema.senderAccounts.id, job.senderAccountId)).limit(1);
  if (!sender || sender.disabled) return stop('The sender account for this job is disabled. An administrator needs to enable it before the send can continue.');
  if (user.role !== 'admin') {
    try {
      await assertSenderAccess(user, job.senderAccountId);
    } catch (err) {
      if (!(err instanceof ApiError)) throw err;
      return stop(`You no longer have permission to send from ${sender.email}. Ask an administrator to restore your access, or cancel this send.`);
    }
  }

  const [moved] = await db.update(schema.sendJobs)
    .set({ status: 'sending', startedAt: sql`coalesce(${schema.sendJobs.startedAt}, now())`, finishedAt: null })
    .where(and(eq(schema.sendJobs.id, jobId), inArray(schema.sendJobs.status, ['queued', 'sending', 'paused'])))
    .returning();
  if (!moved) return { job: await jobView(jobId), processed: [], done: true };
  if (!job.startedAt) await audit(c, 'job.start', { targetType: 'job', targetId: jobId, meta: { senderEmail: sender.email } });

  let creds: { user: string; pass: string };
  try { creds = senderCredentials(sender); } catch (err) { return stop((err as Error).message); }

  let attachments: MailAttachment[];
  try {
    const ids = (await db.select({ id: schema.jobAttachments.uploadId }).from(schema.jobAttachments).where(eq(schema.jobAttachments.jobId, jobId))).map(r => r.id);
    attachments = await resolveAttachments(ids, job.userId);
  } catch (err) {
    return stop(`Attachment problem: ${(err as Error).message}`);
  }
  const attachmentBytes = attachments.reduce((s, a) => s + a.content.length, 0);

  const content = contentOf(job);
  const images = new ImageCache();
  let transport: Transport | null = null;
  let error: string | undefined; let quotaExhausted = false;

  try {
    // Only claim while enough of the 60 s function budget remains to finish the send safely.
    while (Date.now() < claimUntil) {
      if (await jobStatus(jobId) !== 'sending') break;

      const msg = await claimNext(jobId);
      if (!msg) {
        // Nothing queued. If another runner (a second tab, an earlier request still finishing)
        // has messages in flight, wait for them here instead of returning at once, so the
        // client does not flood /run while they finish.
        if (!(await inFlightCount(jobId))) break;
        const wait = inFlightWaitMs(Date.now(), claimUntil);
        if (wait <= 0) break;
        await sleep(wait);
        continue;
      }

      // Paused or cancelled between the status check and the claim: hand the message back
      // (or skip it, if cancelled) without sending.
      if (await jobStatus(jobId) !== 'sending') {
        processed.push((await settleUnsent(jobId, msg, { message: msg.error, retryable: false, accountLevel: true })).row);
        break;
      }

      // Reserve this message's recipients against today's quota before sending (atomic).
      const reservedDay = await reserveQuota(sender.id, msg.recipientCount, sender.dailyLimit);
      if (!reservedDay) {
        const used = await usedToday(sender.id);
        error = `Today's limit of ${fmt(sender.dailyLimit)} recipients for ${sender.email} is reached (${fmt(used)} used). The send is paused; resume it tomorrow or continue from another account.`;
        processed.push((await settleUnsent(jobId, msg, { message: error, retryable: false, accountLevel: true })).row);
        quotaExhausted = true; await pauseJob(jobId);
        break;
      }

      let result: SendOutcome;
      try {
        // Render exactly as the preview does, with this message's row.
        const to = msg.to;
        let unsubUrl: string | null = null;
        if (job.unsubscribeEnabled) {
          const first = firstAddress(to);
          unsubUrl = job.mode === 'personal' && first
            ? unsubscribeUrl(first, sender.id, jobId)
            : `mailto:${sender.email}?subject=unsubscribe`;
        }
        const rendered = renderEmail({ content, sender, row: (msg.rowData || {}) as Record<string, string>, unsubscribeUrl: unsubUrl });
        const inlined = await images.inline(rendered.html, rendered.hostedImageIds);
        const imageBytes = inlined.attachments.reduce((s, a) => s + a.content.length, 0);
        const estimated = estimateMessageBytes(inlined.html, rendered.text, attachmentBytes + imageBytes);

        if (estimated > GMAIL_MAX_MESSAGE_BYTES) {
          result = { ok: false, retryable: false, kind: 'size', message: `This email would be about ${(estimated / 1024 / 1024).toFixed(1)} MB once encoded, over Gmail's 25 MB limit. Remove or shrink attachments.` };
        } else {
          const opts = buildMailOptions({
            sender, to, cc: msg.cc, bcc: msg.bcc, subject: rendered.subject, html: inlined.html, text: rendered.text,
            attachments: [...attachments, ...inlined.attachments],
            unsubscribe: job.unsubscribeEnabled, unsubscribeUrl: unsubUrl,
          });
          try {
            transport ??= createTransport(creds.user, creds.pass);
            const info = await withTimeout(transport.sendMail(opts), SEND_TIMEOUT_MS);
            result = { ok: true, messageId: sentMessageId(info, opts) };
          } catch (err) {
            if (err instanceof SendTimeoutError) {
              result = { ok: false, retryable: true, kind: 'timeout', message: `${err.message} It will be retried automatically.` };
            } else {
              const cl = classifyError(err);
              result = { ok: false, message: cl.message, retryable: cl.retryable, kind: cl.kind };
            }
          }
        }
      } catch (err) {
        // Rendering or loading images failed before anything reached Gmail.
        result = { ok: false, retryable: true, kind: 'prepare', message: `This email could not be prepared: ${(err as Error).message}` };
      }

      if (result.ok) {
        const messageId = result.messageId;
        try {
          processed.push(await retryAsync(() => recordSent(jobId, msg, messageId), RECORD_SENT_BACKOFF_MS, {
            onError: (err, attempt) => console.error(`[jobs] job ${jobId} message ${msg.id}: recording the send failed (attempt ${attempt} of ${RECORD_SENT_BACKOFF_MS.length + 1}):`, (err as Error).message),
          }));
        } catch (err) {
          // Gmail has the message. Never re-queue it: it stays 'sending' and the next run's
          // stale sweep turns it into "Delivery status unknown" (failed), never a resend.
          console.error(`[jobs] CRITICAL: job ${jobId} message ${msg.id} (#${msg.seq} to ${msg.to}) was ACCEPTED by Gmail (Message-ID ${messageId ?? 'unknown'}) but could not be recorded as sent. It must NOT be resent.`, err);
          error = `Email #${msg.seq} to ${msg.to} was sent, but the server could not record it because the database did not respond. The send is paused so nothing goes out twice. That email will later show as failed with "Delivery status unknown"; do not retry it, it was delivered.`;
          processed.push({ ...msg, error });
          try { await pauseJob(jobId); } catch { /* database unreachable; the response carries the error */ }
          break;
        }
      } else {
        await refundQuota(sender.id, reservedDay, msg.recipientCount);
        const accountLevel = result.kind === 'auth' || result.kind === 'quota';
        const { row, action } = await settleUnsent(jobId, msg, { message: result.message, retryable: result.retryable, accountLevel });
        processed.push(row);
        if (result.kind === 'network' || result.kind === 'timeout') { transport?.close(); transport = null; }
        if (action === 'release') {
          // Account-level problem: the message stays queued for when it is fixed, and the job pauses.
          error = result.message;
          if (result.kind === 'quota') quotaExhausted = true;
          await pauseJob(jobId);
          break;
        }
        if (action === 'skip') break; // cancelled while this message was in flight
        if (action === 'requeue') {
          // Back off before the retry (ported from the v1 client: 2 s, then 6 s).
          await sleep(Math.min(msg.attempts === 1 ? 2000 : 6000, claimUntil - Date.now()));
          continue;
        }
      }

      // Pace sends with jitter so Gmail sees a human-like rate.
      const delay = paceDelayMs();
      if (Date.now() + delay >= claimUntil) break;
      await sleep(delay);
    }
  } finally {
    transport?.close();
  }

  // Finish when nothing is left to send.
  const [left] = await db.select({ n: sql<number>`count(*)::int` }).from(schema.sendMessages)
    .where(and(eq(schema.sendMessages.jobId, jobId), inArray(schema.sendMessages.status, ['queued', 'sending'])));
  if ((left?.n ?? 0) === 0) {
    const [done] = await db.update(schema.sendJobs).set({ status: 'completed', finishedAt: new Date() })
      .where(and(eq(schema.sendJobs.id, jobId), eq(schema.sendJobs.status, 'sending'))).returning();
    if (done) {
      await audit(c, 'job.complete', { targetType: 'job', targetId: jobId, meta: { sent: done.sentCount, failed: done.failedCount, skipped: done.skippedCount } });
    }
  }

  const view = await jobView(jobId);
  return {
    job: view, processed: processed.map(toJobMessage),
    done: view.status !== 'sending' && view.status !== 'queued',
    ...(quotaExhausted ? { quotaExhausted: true } : {}),
    ...(error ? { error } : {}),
  };
}

// ---------------------------------------------------------------- state changes
export async function pause(c: Context<AppEnv>, user: UserRow, id: string): Promise<Job> {
  const { job } = await getJobForUser(user, id);
  if (job.status === 'queued' || job.status === 'sending') {
    await pauseJob(id);
    await audit(c, 'job.pause', { targetType: 'job', targetId: id });
  }
  return jobView(id);
}

export async function cancel(c: Context<AppEnv>, user: UserRow, id: string): Promise<Job> {
  const { job } = await getJobForUser(user, id);
  if (job.status === 'completed' || job.status === 'cancelled') return jobView(id);
  await db.transaction(async (tx) => {
    // Locks the job row first; settleUnsent does the same, so an in-flight message that then
    // fails retryably is skipped (not re-queued) and the counters add up.
    await tx.update(schema.sendJobs).set({ status: 'cancelled', finishedAt: new Date() }).where(eq(schema.sendJobs.id, id));
    const skipped = await tx.update(schema.sendMessages).set({ status: 'skipped', error: CANCELLED_ERROR })
      .where(and(eq(schema.sendMessages.jobId, id), eq(schema.sendMessages.status, 'queued')))
      .returning({ id: schema.sendMessages.id });
    if (skipped.length) {
      await tx.update(schema.sendJobs).set({ skippedCount: sql`${schema.sendJobs.skippedCount} + ${skipped.length}` }).where(eq(schema.sendJobs.id, id));
    }
  });
  // Messages whose runner died mid-send may have been delivered: "Delivery status unknown", not "skipped".
  await sweepStaleClaims(id);
  await audit(c, 'job.cancel', { targetType: 'job', targetId: id });
  return jobView(id);
}

export async function retryFailed(c: Context<AppEnv>, user: UserRow, id: string): Promise<Job> {
  const { job } = await getJobForUser(user, id);
  if (job.status === 'sending') throw badRequest('This send is still running. Pause it first, or wait for it to finish.');
  const n = await db.transaction(async (tx) => {
    const rows = await tx.update(schema.sendMessages).set({ status: 'queued', attempts: 0, error: null, retryable: null, claimedAt: null })
      .where(and(eq(schema.sendMessages.jobId, id), eq(schema.sendMessages.status, 'failed')))
      .returning({ id: schema.sendMessages.id });
    if (rows.length) {
      await tx.update(schema.sendJobs).set({
        status: 'queued', finishedAt: null,
        failedCount: sql`greatest(${schema.sendJobs.failedCount} - ${rows.length}, 0)`,
      }).where(eq(schema.sendJobs.id, id));
    }
    return rows.length;
  });
  if (!n) throw badRequest('There are no failed emails to retry in this send.');
  await audit(c, 'job.retry_failed', { targetType: 'job', targetId: id, meta: { requeued: n } });
  return jobView(id);
}

// ---------------------------------------------------------------- queries
const PAGE = 30;

export async function listJobs(user: UserRow, scope: 'mine' | 'all', cursor?: string): Promise<{ jobs: Job[]; nextCursor: string | null }> {
  if (scope === 'all' && user.role !== 'admin') throw forbidden('Only administrators can see everyone\'s sends.');
  const conds = [];
  if (scope === 'mine') conds.push(eq(schema.sendJobs.userId, user.id));
  if (cursor) {
    const d = new Date(cursor);
    if (Number.isNaN(d.getTime())) throw badRequest('Invalid cursor.');
    conds.push(lt(schema.sendJobs.createdAt, d));
  }
  const rows = await db.select({ job: schema.sendJobs, senderEmail: schema.senderAccounts.email, userName: schema.users.name })
    .from(schema.sendJobs)
    .innerJoin(schema.senderAccounts, eq(schema.senderAccounts.id, schema.sendJobs.senderAccountId))
    .innerJoin(schema.users, eq(schema.users.id, schema.sendJobs.userId))
    .where(conds.length ? and(...conds) : undefined)
    .orderBy(desc(schema.sendJobs.createdAt))
    .limit(PAGE + 1);
  const page = rows.slice(0, PAGE);
  return {
    jobs: page.map(r => toJob(r.job, r.senderEmail, r.userName)),
    nextCursor: rows.length > PAGE ? page[page.length - 1].job.createdAt.toISOString() : null,
  };
}

export async function activeJobs(user: UserRow): Promise<Job[]> {
  const rows = await db.select({ job: schema.sendJobs, senderEmail: schema.senderAccounts.email, userName: schema.users.name })
    .from(schema.sendJobs)
    .innerJoin(schema.senderAccounts, eq(schema.senderAccounts.id, schema.sendJobs.senderAccountId))
    .innerJoin(schema.users, eq(schema.users.id, schema.sendJobs.userId))
    .where(and(eq(schema.sendJobs.userId, user.id), inArray(schema.sendJobs.status, ['queued', 'sending', 'paused'])))
    .orderBy(desc(schema.sendJobs.createdAt));
  return rows.map(r => toJob(r.job, r.senderEmail, r.userName));
}

export async function jobDetail(user: UserRow, id: string): Promise<{ job: Job; messages: JobMessage[] }> {
  const r = await getJobForUser(user, id);
  const msgs = await db.select().from(schema.sendMessages).where(eq(schema.sendMessages.jobId, id)).orderBy(schema.sendMessages.seq);
  return { job: toJob(r.job, r.senderEmail, r.userName), messages: msgs.map(toJobMessage) };
}

export async function jobMessagesForExport(user: UserRow, id: string) {
  const r = await getJobForUser(user, id);
  const msgs = await db.select().from(schema.sendMessages).where(eq(schema.sendMessages.jobId, id)).orderBy(schema.sendMessages.seq);
  return { job: r.job, messages: msgs };
}

// ---------------------------------------------------------------- test send
export async function testSend(c: Context<AppEnv>, user: UserRow, req: TestSendRequest): Promise<{ results: { to: string; success: boolean; error?: string }[] }> {
  const sender: SenderRow = await assertSenderAccess(user, req.senderAccountId);
  const to = uniqueEmails(...req.to);

  let creds: { user: string; pass: string };
  try { creds = senderCredentials(sender); } catch (err) { throw badRequest((err as Error).message); }
  let attachments: MailAttachment[];
  try { attachments = await resolveAttachments(req.attachmentIds || [], user.id); } catch (err) { throw badRequest(`Attachment problem: ${(err as Error).message}`); }

  const rendered = renderEmail({ content: req, sender, row: req.row || {}, unsubscribeUrl: null });
  const inlined = await new ImageCache().inline(rendered.html, rendered.hostedImageIds);
  const subject = '[Test] ' + rendered.subject;

  // Same atomic reservation as the runner; whatever is not sent is refunded below.
  const reservedDay = await reserveQuota(sender.id, to.length, sender.dailyLimit);
  if (!reservedDay) {
    const q = checkQuota(to.length, sender.dailyLimit, await usedToday(sender.id), sender.email);
    throw quotaExhausted(q.message ?? `Today's limit of ${fmt(sender.dailyLimit)} recipients for ${sender.email} is reached. Try again tomorrow or use another account.`, { needed: to.length, remaining: q.remaining });
  }

  const results: { to: string; success: boolean; error?: string }[] = [];
  let sent = 0;
  let transport: Transport | null = null;
  try {
    transport = createTransport(creds.user, creds.pass);
    for (const addr of to) {
      const opts = buildMailOptions({ sender, to: addr, subject, html: inlined.html, text: rendered.text, attachments: [...attachments, ...inlined.attachments] });
      try {
        await withTimeout(transport.sendMail(opts), SEND_TIMEOUT_MS);
        results.push({ to: addr, success: true }); sent++;
      } catch (err) {
        const timedOut = err instanceof SendTimeoutError;
        const cl = timedOut ? { kind: 'timeout', message: (err as Error).message } : classifyError(err);
        results.push({ to: addr, success: false, error: cl.message });
        // Account-level problem, or a hung connection (stay inside the request time limit): stop here.
        if (timedOut || cl.kind === 'auth' || cl.kind === 'quota') {
          const why = timedOut ? 'Not sent: an earlier test email timed out.' : cl.message;
          for (const rest of to.slice(results.length)) results.push({ to: rest, success: false, error: why });
          break;
        }
      }
    }
  } finally {
    transport?.close();
    await refundQuota(sender.id, reservedDay, to.length - sent);
  }
  await audit(c, 'job.test_send', { targetType: 'sender', targetId: sender.id, meta: { senderEmail: sender.email, to, sent } });
  return { results };
}
