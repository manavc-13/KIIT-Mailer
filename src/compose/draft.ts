// Draft autosave: the compose work in progress is saved to localStorage, debounced and keyed
// per user (`km_draft_<userId>`), and restored once per sign-in with a "Start fresh" action.
// Large CSV row sets (over ~1 MB serialized) are not persisted; the user is asked to re-upload.
import { useEffect } from 'react';
import { toast } from 'sonner';
import { useComposeStore, initialData, type ComposeData, type ComposeState } from './store';

export const DRAFT_PREFIX = 'km_draft_';
const DRAFT_VERSION = 1;
const MAX_CSV_BYTES = 1_000_000;
const SAVE_DEBOUNCE_MS = 800;

const draftKey = (userId: string) => DRAFT_PREFIX + userId;

type DraftPayload = Omit<ComposeData, 'attachments'> & {
  v: number; savedAt: string;
  attachments: ComposeData['attachments'];
  csvDropped?: boolean;
};

const PERSISTED_KEYS = Object.keys(initialData()) as (keyof ComposeData)[];

export function serializeDraft(s: ComposeState): string {
  const data: Partial<DraftPayload> = { v: DRAFT_VERSION, savedAt: new Date().toISOString() };
  for (const k of PERSISTED_KEYS) (data as Record<string, unknown>)[k] = s[k];
  // Only finished uploads survive a reload; in-flight ones can't be resumed.
  data.attachments = s.attachments.filter(a => a.status === 'ready' && a.ref);
  if (s.csvRows) {
    const csvJson = JSON.stringify(s.csvRows);
    if (csvJson.length > MAX_CSV_BYTES) { data.csvRows = null; data.csvDropped = true; }
  }
  return JSON.stringify(data);
}

function isWorthRestoring(d: Partial<DraftPayload>): boolean {
  const body = String(d.bodyHtml ?? '').replace(/<[^>]+>/g, '').trim();
  const gridHasData = (d.gridRows ?? []).some(r => Object.values(r).some(v => String(v ?? '').trim()));
  return !!(String(d.subject ?? '').trim() || body || String(d.singleEmail ?? '').trim() || d.csvFileName || gridHasData
    || (d.attachments ?? []).length);
}

export function readDraft(userId: string): Partial<DraftPayload> | null {
  try {
    const raw = localStorage.getItem(draftKey(userId));
    if (!raw) return null;
    const d = JSON.parse(raw) as Partial<DraftPayload>;
    if (!d || d.v !== DRAFT_VERSION) return null;
    return isWorthRestoring(d) ? d : null;
  } catch {
    return null;
  }
}

export function clearDraft(userId: string | null | undefined) {
  if (!userId) return;
  try { localStorage.removeItem(draftKey(userId)); } catch { /* storage unavailable */ }
}

/**
 * Restores the draft once for this user, then autosaves changes. Mount once (ComposePage).
 * A different user signing in on the same tab gets a clean store.
 */
export function useDraftAutosave(userId: string | null | undefined) {
  useEffect(() => {
    if (!userId) return;
    const store = useComposeStore.getState();
    if (store.ownerId !== userId) {
      store.reset(userId);
      const d = readDraft(userId);
      if (d) {
        const patch: Partial<ComposeData> = {};
        for (const k of PERSISTED_KEYS) if (k in d) (patch as Record<string, unknown>)[k] = (d as Record<string, unknown>)[k];
        const csvGone = !!d.csvDropped || (d.source === 'csv' && !d.csvRows && !!d.csvFileName);
        useComposeStore.setState(s => ({
          ...patch,
          csvRows: csvGone ? null : patch.csvRows ?? null,
          csvNeedsReupload: csvGone,
          editorKey: s.editorKey + 1,
          dirty: true,
        }));
        toast('Restored your draft', {
          description: csvGone
            ? `The CSV "${d.csvFileName ?? 'file'}" was too large to keep; re-upload it to continue.`
            : d.savedAt ? `Last edited ${new Date(d.savedAt).toLocaleString()}` : undefined,
          duration: 8000,
          action: {
            label: 'Start fresh',
            onClick: () => { clearDraft(userId); useComposeStore.getState().reset(userId); },
          },
        });
      }
    }

    let timer: ReturnType<typeof setTimeout> | null = null;
    const unsub = useComposeStore.subscribe((s, prev) => {
      if (s.ownerId !== userId) return;
      // Ignore session-only changes (preview navigation etc.).
      if (PERSISTED_KEYS.every(k => s[k] === prev[k])) return;
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => {
        try { localStorage.setItem(draftKey(userId), serializeDraft(useComposeStore.getState())); }
        catch { /* quota exceeded: the draft just won't persist */ }
      }, SAVE_DEBOUNCE_MS);
    });
    return () => {
      unsub();
      if (timer) {
        clearTimeout(timer);
        try { localStorage.setItem(draftKey(userId), serializeDraft(useComposeStore.getState())); } catch { /* ignore */ }
      }
    };
  }, [userId]);
}
