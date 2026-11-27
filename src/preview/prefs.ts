// Per-browser preview preferences (selected client tab, compare mode, checks panel).
const PREFIX = 'kiit-mailer.preview.';

export function readPref(key: string, fallback: string): string {
  try { return window.localStorage.getItem(PREFIX + key) ?? fallback; } catch { return fallback; }
}

export function writePref(key: string, value: string): void {
  try { window.localStorage.setItem(PREFIX + key, value); } catch { /* storage unavailable */ }
}
