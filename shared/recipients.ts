// Recipient list logic shared by the Compose UI and (optionally) the server.
// Ported from legacy/public/js/recipients.js: column-role auto-detection,
// per-row CC/BCC merge rules, validation and job-recipient building.
//
// CC/BCC merge rules (unchanged from v1, see README "CC/BCC columns"):
//   final CC  = row CC  ∪ global CC   (de-duplicated, minus the To address)
//   final BCC = row BCC ∪ global BCC  (de-duplicated, minus To and the final CC)
import type { JobMode, JobRecipient } from './types.js';
import {
  parseEmailList, invalidEmails, uniqueEmails, isValidEmail, hasValidEmailList, substitute,
  normalizeEmailList, MAX_RECIPIENTS_PER_MESSAGE,
} from './text.js';

export type RecipientRole = 'email' | 'name' | 'cc' | 'bcc';
export type RoleMap = Record<RecipientRole, string | null>;
export type RecipientRow = Record<string, string>;

export const RECIPIENT_ROLES: RecipientRole[] = ['email', 'name', 'cc', 'bcc'];
export const ROLE_LABELS: Record<RecipientRole, string> = { email: 'Email', name: 'Name', cc: 'CC', bcc: 'BCC' };

/** Column-role auto-detection aliases, compared against the lower-cased, BOM/whitespace-trimmed header. */
export const ROLE_ALIASES: Record<RecipientRole, string[]> = {
  email: ['email', 'e-mail', 'email address', 'e-mail address', 'mail', 'emailid', 'email id', 'email-id', 'mail id', 'to'],
  name: ['name', 'full name', 'recipient', 'recipient name', 'fullname', 'full-name'],
  cc: ['cc', 'cc email', 'copy', 'cc emails', 'cc address', 'cc addresses'],
  bcc: ['bcc', 'bcc email', 'blind copy', 'bcc emails', 'bcc address', 'bcc addresses'],
};

export const EMPTY_ROLE_MAP: RoleMap = { email: null, name: null, cc: null, bcc: null };
/** Role map for the single-recipient form (fixed Name / Email fields). */
export const SINGLE_ROLE_MAP: RoleMap = { email: 'Email', name: 'Name', cc: null, bcc: null };
/** Default columns of a fresh manual grid. */
export const DEFAULT_GRID_COLUMNS = ['Name', 'Email', 'CC', 'BCC'];
export const DEFAULT_GRID_ROLE_MAP: RoleMap = { email: 'Email', name: 'Name', cc: 'CC', bcc: 'BCC' };

export function normalizeHeader(h: unknown): string {
  return String(h ?? '').replace(/^﻿/, '').trim().toLowerCase().replace(/\s+/g, ' ');
}

/** Auto-detect which column plays which role. First matching column wins; a column is never given two roles. */
export function detectRoleMap(headers: string[]): RoleMap {
  const map: RoleMap = { ...EMPTY_ROLE_MAP };
  const used = new Set<string>();
  const normalized = headers.map(raw => ({ raw, norm: normalizeHeader(raw) }));
  for (const role of RECIPIENT_ROLES) {
    const aliases = ROLE_ALIASES[role];
    // Prefer alias order (so "Email" beats "To" when both exist), then column order.
    for (const alias of aliases) {
      const hit = normalized.find(h => h.norm === alias && !used.has(h.raw));
      if (hit) { map[role] = hit.raw; used.add(hit.raw); break; }
    }
  }
  return map;
}

/** Drop role assignments that point at columns that no longer exist. */
export function sanitizeRoleMap(roleMap: RoleMap, headers: string[]): RoleMap {
  const set = new Set(headers);
  const out: RoleMap = { ...EMPTY_ROLE_MAP };
  for (const role of RECIPIENT_ROLES) out[role] = roleMap[role] && set.has(roleMap[role]!) ? roleMap[role] : null;
  return out;
}

/**
 * Copy of the row with canonical `Email` / `Name` keys added, so {Email} and {Name}
 * always work as placeholders even when the source column is named differently.
 */
export function canonicalizeRow(row: RecipientRow, roleMap: RoleMap): RecipientRow {
  const out: RecipientRow = { ...row };
  if (roleMap.email && row[roleMap.email] != null) out.Email = normalizeEmailList(row[roleMap.email]);
  if (roleMap.name && row[roleMap.name] != null) out.Name = String(row[roleMap.name]).trim();
  return out;
}

export function getEmailFromRow(row: RecipientRow, roleMap: RoleMap): string {
  const key = roleMap.email || 'Email';
  return String(row[key] ?? '').trim();
}

export function isBlankRow(row: RecipientRow): boolean {
  return !Object.values(row).some(v => String(v ?? '').trim());
}

export function nonEmptyRows(rows: RecipientRow[]): RecipientRow[] {
  return rows.filter(r => !isBlankRow(r));
}

export interface ResolvedCopies { cc: string; bcc: string; invalid: string[] }

/**
 * Merge the row's CC/BCC with the global CC/BCC. Case-insensitive de-dupe (first spelling wins),
 * never repeats a To address, and BCC never repeats a CC address. Global fields may use
 * placeholders (e.g. `{Supervisor}`), resolved against the row. Invalid addresses are dropped
 * and reported in `invalid`.
 */
export function resolveCopies(
  row: RecipientRow, roleMap: RoleMap, globalCcRaw: string, globalBccRaw: string, toEmail: string,
): ResolvedCopies {
  const globalCc = substitute(globalCcRaw || '', row);
  const globalBcc = substitute(globalBccRaw || '', row);
  const rowCc = roleMap.cc ? String(row[roleMap.cc] ?? '') : '';
  const rowBcc = roleMap.bcc ? String(row[roleMap.bcc] ?? '') : '';

  const toSet = new Set(parseEmailList(toEmail).map(e => e.toLowerCase()));
  const invalid = [...new Set([
    ...invalidEmails(rowCc), ...invalidEmails(rowBcc), ...invalidEmails(globalCc), ...invalidEmails(globalBcc),
  ])];

  const cc = uniqueEmails(rowCc, globalCc).filter(e => isValidEmail(e) && !toSet.has(e.toLowerCase()));
  const ccSet = new Set(cc.map(e => e.toLowerCase()));
  const bcc = uniqueEmails(rowBcc, globalBcc)
    .filter(e => isValidEmail(e) && !toSet.has(e.toLowerCase()) && !ccSet.has(e.toLowerCase()));

  return { cc: cc.join(', '), bcc: bcc.join(', '), invalid };
}

export interface ValidationStats {
  /** Non-blank rows. */
  total: number;
  /** Rows whose email cell is a valid address (or list of addresses). */
  valid: number;
  /** Distinct invalid email cells. */
  invalid: string[];
  /** Distinct addresses that appear more than once (lower-cased). */
  duplicates: string[];
  /** Extra rows that repeat an earlier address (what dedupe would drop). */
  duplicateRows: number;
  /** Rows with no email at all. */
  missingEmail: number;
  /** Rows with no name (all rows when no Name column is mapped). */
  missingName: number;
  uniqueCount: number;
}

export function computeValidation(rowsIn: RecipientRow[], roleMap: RoleMap): ValidationStats {
  const rows = nonEmptyRows(rowsIn);
  const emails = rows.map(r => getEmailFromRow(r, roleMap));
  const present = emails.filter(Boolean);
  const invalid = [...new Set(present.filter(e => !hasValidEmailList(e)))];
  const validEmails = present.filter(e => hasValidEmailList(e)).map(e => normalizeEmailList(e).toLowerCase());
  const seen = new Set<string>();
  const dupes = new Set<string>();
  let duplicateRows = 0;
  for (const e of validEmails) {
    if (seen.has(e)) { dupes.add(e); duplicateRows++; } else seen.add(e);
  }
  const missingName = roleMap.name
    ? rows.filter(r => !String(r[roleMap.name!] ?? '').trim()).length
    : rows.length;
  return {
    total: rows.length,
    valid: validEmails.length,
    invalid,
    duplicates: [...dupes],
    duplicateRows,
    missingEmail: emails.length - present.length,
    missingName,
    uniqueCount: seen.size,
  };
}

/** Keep the first row for each (case-insensitive) email; rows without an email are dropped. */
export function dedupeRows(rows: RecipientRow[], roleMap: RoleMap): { rows: RecipientRow[]; dropped: number } {
  const seen = new Set<string>();
  const out: RecipientRow[] = [];
  let dropped = 0;
  for (const r of rows) {
    const key = normalizeEmailList(getEmailFromRow(r, roleMap)).toLowerCase();
    if (!key) continue;
    if (seen.has(key)) { dropped++; continue; }
    seen.add(key);
    out.push(r);
  }
  return { rows: out, dropped };
}

/**
 * Placeholder names offered in the editor: every column, plus the canonical `Name` / `Email`
 * aliases when those roles are mapped to differently-named columns.
 */
export function placeholderHeaders(headers: string[], roleMap: RoleMap): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  const add = (h: string) => {
    const t = String(h ?? '').replace(/^﻿/, '').trim();
    if (!t || seen.has(t.toLowerCase())) return;
    seen.add(t.toLowerCase()); out.push(t);
  };
  if (roleMap.name) add('Name');
  if (roleMap.email) add('Email');
  headers.forEach(add);
  return out;
}

export function chunk<T>(arr: T[], size: number): T[][] {
  const n = Math.max(1, Math.floor(size));
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
  return out;
}

/** Number of SMTP messages for a BCC broadcast to `n` addresses (Gmail caps recipients per message). */
export function bccMessageCount(n: number, perMessage = MAX_RECIPIENTS_PER_MESSAGE): number {
  return n <= 0 ? 0 : Math.ceil(n / perMessage);
}

export interface BuildJobRecipientsInput {
  rows: RecipientRow[];
  roleMap: RoleMap;
  mode: JobMode;
  globalCc?: string;
  globalBcc?: string;
  /** Drop repeated To addresses (always on for BCC broadcast). */
  dedupe?: boolean;
}

export interface BuildJobRecipientsResult {
  recipients: JobRecipient[];
  /** SMTP messages the job will send (personal: one per recipient; broadcast: BCC chunks). */
  messageCount: number;
  /** Distinct primary recipients (To addresses in personal mode, BCC addresses in broadcast mode). */
  recipientCount: number;
  /** Total CC / BCC copies across all messages. */
  ccTotal: number;
  bccTotal: number;
  /** To + CC + BCC across all messages: what Gmail counts against the daily quota. */
  quotaCost: number;
  /** Email cells that were not valid addresses (these rows are skipped). */
  skippedInvalid: string[];
  /** Non-blank rows without an email (skipped). */
  skippedMissing: number;
  droppedDuplicates: number;
  /** Invalid CC/BCC addresses that were dropped from copies. */
  invalidCopies: string[];
  /** Per-row CC/BCC columns are ignored in BCC broadcast mode; true when any row had them. */
  ignoredRowCopies: boolean;
}

/**
 * Turn the recipient rows into the `recipients` array of a CreateJobRequest.
 *
 * - personal: one entry per valid row. `cc` / `bcc` are fully resolved (row ∪ global, merge rules
 *   above), and `data` is the canonicalized row for placeholder substitution.
 * - bcc_broadcast: one entry per distinct valid address (`to` = that BCC address, no data, since
 *   only global placeholders work). The server chunks them into messages of <= 90. Global CC/BCC
 *   are sent separately on the request (`globalCc` / `globalBcc`) and apply to every chunk.
 */
export function buildJobRecipients(input: BuildJobRecipientsInput): BuildJobRecipientsResult {
  const { roleMap, mode } = input;
  const globalCc = input.globalCc ?? '';
  const globalBcc = input.globalBcc ?? '';
  const rows = nonEmptyRows(input.rows);

  const skippedInvalid = new Set<string>();
  let skippedMissing = 0;
  let validRows: RecipientRow[] = [];
  for (const r of rows) {
    const email = getEmailFromRow(r, roleMap);
    if (!email) { skippedMissing++; continue; }
    if (!hasValidEmailList(email)) { skippedInvalid.add(email); continue; }
    validRows.push(r);
  }

  let droppedDuplicates = 0;
  if (input.dedupe || mode === 'bcc_broadcast') {
    const d = dedupeRows(validRows, roleMap);
    validRows = d.rows;
    droppedDuplicates = d.dropped;
  }

  const invalidCopies = new Set<string>();
  const base = {
    skippedInvalid: [...skippedInvalid], skippedMissing, droppedDuplicates,
  };

  if (mode === 'bcc_broadcast') {
    // One flat, de-duplicated address list (a row may hold several addresses).
    const addresses = uniqueEmails(...validRows.map(r => getEmailFromRow(r, roleMap)));
    const ignoredRowCopies = validRows.some(r =>
      (roleMap.cc && String(r[roleMap.cc] ?? '').trim()) || (roleMap.bcc && String(r[roleMap.bcc] ?? '').trim()));
    const lower = new Set(addresses.map(a => a.toLowerCase()));
    const gCc = uniqueEmails(globalCc).filter(e => {
      if (!isValidEmail(e)) { invalidCopies.add(e); return false; }
      return !lower.has(e.toLowerCase());
    });
    const ccLower = new Set(gCc.map(e => e.toLowerCase()));
    const gBcc = uniqueEmails(globalBcc).filter(e => {
      if (!isValidEmail(e)) { invalidCopies.add(e); return false; }
      return !lower.has(e.toLowerCase()) && !ccLower.has(e.toLowerCase());
    });
    const messageCount = bccMessageCount(addresses.length);
    // Each chunk: To = the sender itself (1) + global CC + global BCC + its share of the list.
    const ccTotal = messageCount * gCc.length;
    const bccTotal = addresses.length + messageCount * gBcc.length;
    return {
      ...base,
      recipients: addresses.map(to => ({ to })),
      messageCount,
      recipientCount: addresses.length,
      ccTotal,
      bccTotal,
      quotaCost: messageCount + ccTotal + bccTotal,
      invalidCopies: [...invalidCopies],
      ignoredRowCopies: !!ignoredRowCopies,
    };
  }

  let ccTotal = 0;
  let bccTotal = 0;
  let toTotal = 0;
  const recipients: JobRecipient[] = validRows.map(row => {
    const to = normalizeEmailList(getEmailFromRow(row, roleMap));
    const data = canonicalizeRow(row, roleMap);
    const copies = resolveCopies(data, roleMap, globalCc, globalBcc, to);
    copies.invalid.forEach(e => invalidCopies.add(e));
    const ccN = parseEmailList(copies.cc).length;
    const bccN = parseEmailList(copies.bcc).length;
    ccTotal += ccN; bccTotal += bccN; toTotal += parseEmailList(to).length;
    const rec: JobRecipient = { to, data };
    if (copies.cc) rec.cc = copies.cc;
    if (copies.bcc) rec.bcc = copies.bcc;
    return rec;
  });

  return {
    ...base,
    recipients,
    messageCount: recipients.length,
    recipientCount: recipients.length,
    ccTotal,
    bccTotal,
    quotaCost: toTotal + ccTotal + bccTotal,
    invalidCopies: [...invalidCopies],
    ignoredRowCopies: false,
  };
}
