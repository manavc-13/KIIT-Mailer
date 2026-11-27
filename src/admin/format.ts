// Date formatting helpers for the admin screens (en-IN, local time).

const dateTimeFmt = new Intl.DateTimeFormat('en-IN', {
  day: 'numeric', month: 'short', year: 'numeric', hour: 'numeric', minute: '2-digit',
});
const dateFmt = new Intl.DateTimeFormat('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
const rtf = new Intl.RelativeTimeFormat('en', { numeric: 'auto' });

export function formatDateTime(iso: string | null | undefined): string {
  if (!iso) return '';
  return dateTimeFmt.format(new Date(iso));
}

export function formatDate(iso: string | null | undefined): string {
  if (!iso) return '';
  return dateFmt.format(new Date(iso));
}

/** "just now", "5 minutes ago", "yesterday", "3 weeks ago", ... */
export function relativeTime(iso: string | null | undefined, now = Date.now()): string {
  if (!iso) return '';
  const diffSec = Math.round((new Date(iso).getTime() - now) / 1000);
  const abs = Math.abs(diffSec);
  if (abs < 45) return 'just now';
  const units: [Intl.RelativeTimeFormatUnit, number][] = [
    ['minute', 60], ['hour', 3600], ['day', 86400], ['week', 604800], ['month', 2629800], ['year', 31557600],
  ];
  let unit: Intl.RelativeTimeFormatUnit = 'minute';
  let size = 60;
  for (const [u, s] of units) { if (abs >= s * 0.9) { unit = u; size = s; } }
  return rtf.format(Math.round(diffSec / size), unit);
}
