// Pure helpers shared by client and server (ported from legacy/public/js/util.js;
// behaviour preserved, now typed).

export function escapeHtml(s: unknown): string {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

export const KIIT_DOMAIN = 'kiit.ac.in';
const EMAIL_RE = /^[^\s@<>()[\],;:"]+@[^\s@<>()[\],;:"]+\.[^\s@<>()[\],;:"]{2,}$/;

export function isValidEmail(email: string): boolean {
  return EMAIL_RE.test(email.trim());
}

export function isKiitEmail(email: string): boolean {
  return email.trim().toLowerCase().endsWith('@' + KIIT_DOMAIN);
}

export function parseEmailList(value: unknown): string[] {
  return String(value ?? '')
    .split(/[\s,;]+/)
    .map(v => v.trim())
    .filter(Boolean);
}

export function normalizeEmailList(value: unknown): string {
  return parseEmailList(value).join(', ');
}

export function invalidEmails(value: unknown): string[] {
  return parseEmailList(value).filter(e => !isValidEmail(e));
}

export function hasValidEmailList(value: unknown): boolean {
  const emails = parseEmailList(value);
  return emails.length > 0 && invalidEmails(value).length === 0;
}

/** Case-insensitive, de-duplicated union of one or more email-list strings (first spelling wins). */
export function uniqueEmails(...lists: unknown[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const list of lists) {
    for (const e of parseEmailList(list)) {
      const key = e.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(e);
    }
  }
  return out;
}

/**
 * Placeholder substitution: {Key} resolves against a row whose keys may be any case.
 * Exact-case key wins; otherwise first case-insensitive match. Unresolved tokens are left verbatim.
 * When `escape` is true, substituted values are HTML-escaped (use for HTML bodies).
 */
export function substitute(str: string, row: Record<string, unknown> | null | undefined, escape = false): string {
  if (!str) return str;
  const lowerIndex: Record<string, unknown> = {};
  for (const k of Object.keys(row || {})) {
    const lk = k.trim().toLowerCase();
    if (!(lk in lowerIndex)) lowerIndex[lk] = (row as Record<string, unknown>)[k];
  }
  return String(str).replace(/\{([^{}<>\n]+)\}/g, (m, key: string) => {
    const trimmed = key.trim();
    let v: unknown;
    if (row && row[trimmed] != null) v = row[trimmed];
    else if (lowerIndex[trimmed.toLowerCase()] != null) v = lowerIndex[trimmed.toLowerCase()];
    else return m;
    return escape ? escapeHtml(v) : String(v);
  });
}

/** All distinct {Placeholder} names used in the given strings. */
export function placeholdersIn(...strs: (string | null | undefined)[]): string[] {
  const set = new Set<string>();
  for (const s of strs) {
    for (const m of String(s ?? '').matchAll(/\{([^{}<>\n]+)\}/g)) set.add(m[1].trim());
  }
  return [...set];
}

export function textToHtml(text: string): string {
  return escapeHtml(text).split(/\n/).map(l => l || '&nbsp;').join('<br>');
}

/** Ensure outgoing HTML opts out of mail-client dark-mode re-theming. Idempotent. */
export function lockLightMode(html: string): string {
  const hasColorScheme = /<meta[^>]+name=["']color-scheme["']/i.test(html);
  const hasSupported = /<meta[^>]+name=["']supported-color-schemes["']/i.test(html);
  const hasLockStyle = /color-scheme\s*:\s*light\s*only/i.test(html);
  const inject =
    (hasColorScheme ? '' : '<meta name="color-scheme" content="light only">') +
    (hasSupported ? '' : '<meta name="supported-color-schemes" content="light">') +
    (hasLockStyle ? '' : '<style>:root{color-scheme:light only;supported-color-schemes:light}</style>');
  if (!inject) return html;
  if (/<head[^>]*>/i.test(html)) return html.replace(/<head[^>]*>/i, m => m + inject);
  if (/<html[^>]*>/i.test(html)) return html.replace(/<html[^>]*>/i, m => m + '<head>' + inject + '</head>');
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">${inject}</head><body>${html}</body></html>`;
}

/** Hidden inbox preview text, placed first in <body>. Padded so clients don't pull body text after it. */
export function injectPreheader(html: string, preheader: string): string {
  if (!preheader) return html;
  const pad = '&#847;&zwnj;&nbsp;'.repeat(60);
  const block = `<div style="display:none;max-height:0;max-width:0;overflow:hidden;mso-hide:all;font-size:1px;line-height:1px;color:transparent;opacity:0;">${escapeHtml(preheader)}${pad}</div>`;
  if (/<body[^>]*>/i.test(html)) return html.replace(/<body[^>]*>/i, m => m + block);
  return block + html;
}

/** Tabular paste parser: detects tab-separated or comma-separated. */
export function parseTabular(text: string): { columns: string[]; rows: Record<string, string>[] } {
  const lines = text.replace(/\r/g, '').split('\n').filter(l => l.length > 0);
  if (lines.length === 0) return { columns: [], rows: [] };
  const sep = lines[0].includes('\t') ? '\t' : ',';
  const splitLine = (line: string): string[] => {
    if (sep === '\t') return line.split('\t').map(s => s.trim());
    const result: string[] = []; let cur = ''; let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQ) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') inQ = false;
        else cur += ch;
      } else if (ch === '"') inQ = true;
      else if (ch === ',') { result.push(cur.trim()); cur = ''; }
      else cur += ch;
    }
    result.push(cur.trim());
    return result;
  };
  const headers = splitLine(lines[0]);
  const rows = lines.slice(1).map(line => {
    const cells = splitLine(line);
    const o: Record<string, string> = {};
    headers.forEach((h, i) => { o[h] = cells[i] ?? ''; });
    return o;
  });
  return { columns: headers, rows };
}

export function toCSV(columns: string[], rows: Record<string, unknown>[]): string {
  const esc = (v: unknown) => {
    const s = String(v ?? '');
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return columns.map(esc).join(',') + '\n' + rows.map(r => columns.map(c => esc(r[c])).join(',')).join('\n') + '\n';
}

export function formatBytes(bytes: number): string {
  const n = Number(bytes) || 0;
  if (n < 1024) return n + ' B';
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
  return (n / 1024 / 1024).toFixed(2) + ' MB';
}

export function encodedSize(rawBytes: number): number {
  return Math.ceil(rawBytes / 3) * 4;
}

/** Gmail clips message HTML above ~102 KB ("[Message clipped] View entire message"). */
export const GMAIL_CLIP_BYTES = 102 * 1024;
/** Gmail SMTP accepts at most 100 recipients per message; stay under it. */
export const MAX_RECIPIENTS_PER_MESSAGE = 90;
