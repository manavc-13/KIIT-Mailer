// One-time cleanup of KIIT Mailer v1 data left in this browser's localStorage.
// v1 kept the Gmail app password in plain text under `kiit_mailer_settings`; v2 keeps
// sender credentials encrypted on the server, so that copy should be removed.
// `kiit_mailer_templates` is deliberately left alone: the Templates page imports it.
import { useCallback } from 'react';
import { toast } from 'sonner';

export const LEGACY_SETTINGS_KEY = 'kiit_mailer_settings';
export const LEGACY_LEFTOVER_KEYS = [
  'kiit_mailer_logs',
  'kiit_mailer_pending_queue',
  'kiit_mailer_daily_count',
  'kiit_mailer_draft',
] as const;
/** Set when the user chose "Keep" for the non-sensitive leftovers, so we don't ask again. */
const LEFTOVERS_DISMISSED_KEY = 'kiit_mailer_v2_legacy_leftovers_dismissed';

function storage(): Storage | null {
  try { return window.localStorage; } catch { return null; }
}

function has(key: string): boolean {
  try { return storage()?.getItem(key) != null; } catch { return false; }
}

function remove(keys: readonly string[]): void {
  const s = storage();
  for (const k of keys) { try { s?.removeItem(k); } catch { /* ignore */ } }
}

/** Which v1 keys are present in this browser. */
export function detectLegacyData(): { settings: boolean; leftovers: string[] } {
  return { settings: has(LEGACY_SETTINGS_KEY), leftovers: LEGACY_LEFTOVER_KEYS.filter(has) };
}

/** Shows sonner toasts offering to remove v1 data. Safe to call more than once. */
export function offerLegacyCleanup(): void {
  const { settings, leftovers } = detectLegacyData();

  if (settings) {
    toast.warning('Old saved settings found in this browser', {
      id: 'legacy-settings',
      description:
        'The previous version of KIIT Mailer stored your Gmail app password here in plain text. '
        + 'Sender credentials are now stored securely on the server, so this copy is no longer needed and should be removed.',
      duration: Infinity,
      action: {
        label: 'Remove it',
        onClick: () => {
          remove([LEGACY_SETTINGS_KEY]);
          toast.success('The old settings and saved app password were removed from this browser.');
        },
      },
      cancel: { label: 'Not now', onClick: () => undefined },
    });
  }

  if (leftovers.length && !has(LEFTOVERS_DISMISSED_KEY)) {
    toast('Old KIIT Mailer data is still stored in this browser', {
      id: 'legacy-leftovers',
      description:
        'The send log, pending queue, daily counter and draft from the previous version are no longer used. '
        + 'Send history now lives on the server. Saved templates are kept so you can import them on the Templates page.',
      duration: 20_000,
      action: {
        label: 'Remove',
        onClick: () => {
          remove(leftovers);
          toast.success('Old send log, queue, counter and draft removed.');
        },
      },
      cancel: {
        label: 'Keep',
        onClick: () => { try { storage()?.setItem(LEFTOVERS_DISMISSED_KEY, '1'); } catch { /* ignore */ } },
      },
    });
  }
}

/** Returns a callback to run right after a successful sign-in. */
export function useLegacyCleanup(): () => void {
  return useCallback(() => {
    // Let the route change settle so the toast appears on the destination page.
    setTimeout(offerLegacyCleanup, 400);
  }, []);
}
