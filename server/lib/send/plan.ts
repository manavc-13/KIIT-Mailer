// Pure job planning: turn a CreateJobRequest's recipients into message rows.
// No DB access, so every rule here is unit-tested (tests/send-plan.test.ts).
import type { JobRecipient } from '../../../shared/types.js';
import {
  substitute, parseEmailList, uniqueEmails, isValidEmail, isKiitEmail, MAX_RECIPIENTS_PER_MESSAGE,
} from '../../../shared/text.js';

export interface PlannedMessage {
  to: string; cc: string | null; bcc: string | null;
  recipientCount: number;
  rowData: Record<string, string>;
  /** Set when the message must not be sent (e.g. To is suppressed); inserted as status 'skipped'. */
  skipReason?: string;
}

export interface PlanResult {
  messages: PlannedMessage[];
  /** Invalid addresses (first spelling), for a 400 listing the first few. */
  invalid: string[];
  /** Entries with no To address at all (1-based positions). */
  emptyRows: number[];
  duplicatesRemoved: number;
  suppressedAddresses: string[];
}

const lc = (s: string) => s.trim().toLowerCase();

/**
 * Legacy "CC/BCC columns" rule, per message:
 *   final CC  = row CC ∪ global CC   (deduped, minus every To address)
 *   final BCC = row BCC ∪ global BCC (deduped, minus To and final CC)
 * Global CC/BCC may contain {placeholders} resolved against the row.
 */
export function mergeCopies(
  to: string[], rowCc: string | undefined | null, rowBcc: string | undefined | null,
  globalCc: string | undefined | null, globalBcc: string | undefined | null, row: Record<string, string> = {},
): { cc: string[]; bcc: string[] } {
  const toSet = new Set(to.map(lc));
  const cc = uniqueEmails(rowCc, substitute(globalCc || '', row)).filter(e => !toSet.has(lc(e)));
  const ccSet = new Set(cc.map(lc));
  const bcc = uniqueEmails(rowBcc, substitute(globalBcc || '', row)).filter(e => !toSet.has(lc(e)) && !ccSet.has(lc(e)));
  return { cc, bcc };
}

const join = (list: string[]) => (list.length ? list.join(', ') : null);

/** Remove suppressed addresses from a list. */
function dropSuppressed(list: string[], suppressed: Set<string>, removed: Set<string>): string[] {
  return list.filter(e => {
    if (suppressed.has(lc(e))) { removed.add(lc(e)); return false; }
    return true;
  });
}

/** Personal mode: one message per recipient entry. */
export function planPersonal(
  recipients: JobRecipient[],
  opts: { globalCc?: string | null; globalBcc?: string | null; dedupe?: boolean; suppressed?: Set<string> } = {},
): PlanResult {
  const suppressed = opts.suppressed ?? new Set<string>();
  const invalid = new Set<string>(); const invalidOut: string[] = [];
  const noteInvalid = (list: string[]) => list.forEach(e => { if (!isValidEmail(e) && !invalid.has(lc(e))) { invalid.add(lc(e)); invalidOut.push(e); } });
  const emptyRows: number[] = [];
  const removed = new Set<string>();
  const seenTo = new Set<string>();
  let duplicatesRemoved = 0;
  const messages: PlannedMessage[] = [];

  recipients.forEach((r, i) => {
    const row = r.data || {};
    const toAll = uniqueEmails(r.to);
    if (!toAll.length) { emptyRows.push(i + 1); return; }
    const { cc, bcc } = mergeCopies(toAll, r.cc, r.bcc, opts.globalCc, opts.globalBcc, row);
    noteInvalid(toAll); noteInvalid(cc); noteInvalid(bcc);

    if (opts.dedupe) {
      const key = toAll.map(lc).sort().join(',');
      if (seenTo.has(key)) { duplicatesRemoved++; return; }
      seenTo.add(key);
    }

    const to = dropSuppressed(toAll, suppressed, removed);
    const ccF = dropSuppressed(cc, suppressed, removed);
    const bccF = dropSuppressed(bcc, suppressed, removed);
    if (!to.length) {
      messages.push({
        to: toAll.join(', '), cc: null, bcc: null, recipientCount: 0, rowData: row,
        skipReason: 'Not sent: this address unsubscribed or is on the suppression list.',
      });
      return;
    }
    messages.push({ to: to.join(', '), cc: join(ccF), bcc: join(bccF), recipientCount: to.length + ccF.length + bccF.length, rowData: row });
  });

  return { messages, invalid: invalidOut, emptyRows, duplicatesRemoved, suppressedAddresses: [...removed] };
}

export function chunk<T>(list: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < list.length; i += size) out.push(list.slice(i, i + size));
  return out;
}

/**
 * BCC broadcast: every address from the entries (to/cc/bcc) plus global BCC, deduped,
 * minus the sender's own address and suppressed addresses, chunked into BCC groups of
 * at most `chunkSize` (90) with To = the sender. Global CC (if any) goes on the first
 * message only, so CC'd people get one copy. Only global placeholders apply (row = {}).
 */
export function planBroadcast(
  recipients: JobRecipient[], senderEmail: string,
  opts: { globalCc?: string | null; globalBcc?: string | null; suppressed?: Set<string>; chunkSize?: number } = {},
): PlanResult {
  const suppressed = opts.suppressed ?? new Set<string>();
  const size = Math.max(1, Math.min(opts.chunkSize ?? MAX_RECIPIENTS_PER_MESSAGE, MAX_RECIPIENTS_PER_MESSAGE));
  const self = lc(senderEmail);
  const all = uniqueEmails(...recipients.flatMap(r => [r.to, r.cc, r.bcc]), opts.globalBcc);
  const cc = uniqueEmails(opts.globalCc).filter(e => lc(e) !== self);
  const ccSet = new Set(cc.map(lc));
  const invalidOut = [...all, ...cc].filter(e => !isValidEmail(e));
  const removed = new Set<string>();
  const before = all.filter(e => lc(e) !== self && !ccSet.has(lc(e)));
  const bccAll = dropSuppressed(before, suppressed, removed);
  const ccF = dropSuppressed(cc, suppressed, removed);
  // CC addresses share the first message's copy budget, so CC + BCC <= size everywhere.
  const firstSize = Math.max(1, size - ccF.length);
  const groups: string[][] = [];
  if (bccAll.length) {
    groups.push(bccAll.slice(0, firstSize));
    groups.push(...chunk(bccAll.slice(firstSize), size));
  } else if (ccF.length) {
    groups.push([]);
  }
  const messages: PlannedMessage[] = groups.map((g, i) => {
    const c = i === 0 ? ccF : [];
    return { to: senderEmail, cc: join(c), bcc: join(g), recipientCount: 1 + c.length + g.length, rowData: {} };
  });
  return { messages, invalid: invalidOut, emptyRows: [], duplicatesRemoved: 0, suppressedAddresses: [...removed] };
}

/** Total recipients Gmail counts against the daily quota (To + CC + BCC). */
export function totalRecipients(messages: PlannedMessage[]): number {
  return messages.reduce((s, m) => s + (m.skipReason ? 0 : m.recipientCount), 0);
}

/** Distinct addresses outside kiit.ac.in across all messages. */
export function externalRecipientCount(messages: PlannedMessage[]): number {
  const set = new Set<string>();
  for (const m of messages) {
    if (m.skipReason) continue;
    for (const e of [...parseEmailList(m.to), ...parseEmailList(m.cc), ...parseEmailList(m.bcc)]) {
      if (!isKiitEmail(e)) set.add(lc(e));
    }
  }
  return set.size;
}

/** 'auto' → on when more than 50 recipients are outside kiit.ac.in. */
export function resolveUnsubscribe(setting: boolean | 'auto', messages: PlannedMessage[]): boolean {
  if (setting === 'auto') return externalRecipientCount(messages) > 50;
  return setting;
}

const fmt = (n: number) => n.toLocaleString('en-US');

export interface QuotaCheck { ok: boolean; remaining: number; message: string | null }

/** Does `needed` fit in what remains of today's quota? */
export function checkQuota(needed: number, dailyLimit: number, usedToday: number, senderEmail: string): QuotaCheck {
  const remaining = Math.max(0, dailyLimit - usedToday);
  if (needed <= remaining) return { ok: true, remaining, message: null };
  return {
    ok: false, remaining,
    message: `This send needs ${fmt(needed)} recipient${needed === 1 ? '' : 's'} but only ${fmt(remaining)} remain${remaining === 1 ? 's' : ''} today for ${senderEmail}. Split the list or use another account.`,
  };
}

/** First To address of a message (used for per-recipient unsubscribe tokens). */
export function firstAddress(list: string | null | undefined): string | null {
  return parseEmailList(list)[0] ?? null;
}

/** Random pacing delay between sends (ms). Env SEND_DELAY_MIN_MS / SEND_DELAY_MAX_MS override 1500–3000. */
export function paceDelayMs(rand = Math.random): number {
  const min = Math.max(0, Number(process.env.SEND_DELAY_MIN_MS ?? 1500) || 0);
  const maxRaw = Number(process.env.SEND_DELAY_MAX_MS ?? 3000);
  const max = Math.max(min, Number.isFinite(maxRaw) ? maxRaw : 3000);
  return Math.round(min + rand() * (max - min));
}
