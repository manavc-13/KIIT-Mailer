// Formatting helpers for the Sent history pages.
import type { Job, JobMessage, JobStatus } from '../../shared/types';
import { parseEmailList } from '../../shared/text';

const rtf = typeof Intl !== 'undefined' && 'RelativeTimeFormat' in Intl
  ? new Intl.RelativeTimeFormat(undefined, { numeric: 'auto' })
  : null;

/** "3 minutes ago", "yesterday", "2 weeks ago". */
export function relativeTime(iso: string | null | undefined, now: number = Date.now()): string {
  if (!iso) return '—';
  const t = new Date(iso).getTime();
  if (Number.isNaN(t)) return '—';
  const s = Math.round((t - now) / 1000);
  const abs = Math.abs(s);
  if (abs < 45) return 'just now';
  const units: [Intl.RelativeTimeFormatUnit, number][] = [
    ['minute', 60], ['hour', 3600], ['day', 86400], ['week', 604800], ['month', 2629800], ['year', 31557600],
  ];
  let unit: Intl.RelativeTimeFormatUnit = 'minute';
  let size = 60;
  for (const [u, sec] of units) { if (abs >= sec * 0.9) { unit = u; size = sec; } }
  const value = Math.round(s / size);
  return rtf ? rtf.format(value, unit) : `${Math.abs(value)} ${unit}${Math.abs(value) === 1 ? '' : 's'} ${value < 0 ? 'ago' : 'from now'}`;
}

/** "27 Nov 2026, 2:41 pm" in the viewer's locale. */
export function absoluteTime(iso: string | null | undefined): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '—';
  return d.toLocaleString(undefined, { day: 'numeric', month: 'short', year: 'numeric', hour: 'numeric', minute: '2-digit' });
}

/** Not finished yet (includes paused): for UI such as the "Open live view" and "Cancel send" buttons. */
export const ACTIVE_STATUSES: JobStatus[] = ['queued', 'sending', 'paused'];
export const isActive = (s: JobStatus) => ACTIVE_STATUSES.includes(s);

/** Statuses that change on their own and are worth polling. A paused job only changes when someone acts on it. */
export const POLLING_STATUSES: JobStatus[] = ['queued', 'sending'];
export const shouldPoll = (s: JobStatus) => POLLING_STATUSES.includes(s);

export function pendingCount(job: Job): number {
  return Math.max(0, job.totalMessages - job.sentCount - job.failedCount - job.skippedCount);
}

export function countAddresses(list: string | null | undefined): number {
  return parseEmailList(list).length;
}

export type MessageFilter = 'all' | 'sent' | 'failed' | 'skipped' | 'pending';

export function matchesFilter(m: JobMessage, f: MessageFilter): boolean {
  if (f === 'all') return true;
  if (f === 'pending') return m.status === 'queued' || m.status === 'sending';
  return m.status === f;
}

export function matchesSearch(m: JobMessage, q: string): boolean {
  const needle = q.trim().toLowerCase();
  if (!needle) return true;
  return [m.to, m.cc, m.bcc, m.error, String(m.seq)].some(v => (v ?? '').toLowerCase().includes(needle));
}
