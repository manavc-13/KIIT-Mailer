// Derived views of the compose store: the active recipient rows/headers/role map, validation,
// the job recipients, the render request for previews and the content for saving templates.
import { useMemo } from 'react';
import { useShallow } from 'zustand/react/shallow';
import {
  SINGLE_ROLE_MAP, buildJobRecipients, canonicalizeRow, computeValidation, nonEmptyRows, placeholderHeaders,
  type BuildJobRecipientsResult, type RecipientRow, type RoleMap, type ValidationStats,
} from '../../shared/recipients';
import { normalizeEmailList } from '../../shared/text';
import type { ComposeContent, RenderRequest, SaveTemplateRequest, SenderAccount } from '../../shared/types';
import { useMySenders } from '@/lib/queries';
import { useComposeStore, type ComposeData, type ComposeState } from './store';

export interface RecipientModel {
  rows: RecipientRow[];
  headers: string[];
  roleMap: RoleMap;
  validation: ValidationStats;
  build: BuildJobRecipientsResult;
  /** Placeholder names for the editor / subject. */
  placeholders: string[];
}

type RecipientSlice = Pick<ComposeData,
  'source' | 'singleName' | 'singleEmail' | 'singleExtras' | 'csvRows' | 'csvHeaders' | 'csvRoleMap'
  | 'gridColumns' | 'gridRows' | 'gridRoleMap' | 'globalCc' | 'globalBcc' | 'mode' | 'dedupe'>;

const pickRecipients = (s: ComposeState): RecipientSlice => ({
  source: s.source, singleName: s.singleName, singleEmail: s.singleEmail, singleExtras: s.singleExtras,
  csvRows: s.csvRows, csvHeaders: s.csvHeaders, csvRoleMap: s.csvRoleMap,
  gridColumns: s.gridColumns, gridRows: s.gridRows, gridRoleMap: s.gridRoleMap,
  globalCc: s.globalCc, globalBcc: s.globalBcc, mode: s.mode, dedupe: s.dedupe,
});

export function deriveRecipients(s: RecipientSlice): RecipientModel {
  let rows: RecipientRow[];
  let headers: string[];
  let roleMap: RoleMap;
  if (s.source === 'single') {
    const row: RecipientRow = { Name: s.singleName.trim(), Email: normalizeEmailList(s.singleEmail) };
    s.singleExtras.forEach(f => { if (f.key.trim()) row[f.key.trim()] = f.value; });
    rows = row.Email || row.Name ? [row] : [];
    headers = ['Name', 'Email', ...s.singleExtras.map(f => f.key.trim()).filter(Boolean)];
    roleMap = SINGLE_ROLE_MAP;
  } else if (s.source === 'csv') {
    rows = s.csvRows ?? [];
    headers = s.csvHeaders;
    roleMap = s.csvRoleMap;
  } else {
    rows = nonEmptyRows(s.gridRows);
    headers = s.gridColumns;
    roleMap = s.gridRoleMap;
  }
  const build = buildJobRecipients({
    rows, roleMap, mode: s.mode, globalCc: s.globalCc, globalBcc: s.globalBcc,
    dedupe: s.source !== 'single' && s.dedupe,
  });
  return {
    rows, headers, roleMap,
    validation: computeValidation(rows, roleMap),
    build,
    placeholders: placeholderHeaders(headers, roleMap),
  };
}

export function useRecipientModel(): RecipientModel {
  const slice = useComposeStore(useShallow(pickRecipients));
  return useMemo(() => deriveRecipients(slice), [slice]);
}

/** Rows the preview can step through (valid recipients, canonicalized). */
export function previewRows(model: RecipientModel, mode: ComposeData['mode']): RecipientRow[] {
  if (mode === 'bcc_broadcast') return [];
  return model.build.recipients.map(r => r.data ?? {});
}

export function firstRecipientRow(model: RecipientModel, mode: ComposeData['mode']): RecipientRow | undefined {
  if (mode === 'bcc_broadcast') return undefined;
  const r = model.build.recipients[0];
  return r?.data ?? (model.rows[0] ? canonicalizeRow(model.rows[0], model.roleMap) : undefined);
}

type ContentSlice = Pick<ComposeData,
  'subject' | 'preheader' | 'editorMode' | 'bodyHtml' | 'customText' | 'customTextEnabled' | 'theme' | 'includeSignature'>;

const pickContent = (s: ComposeState): ContentSlice => ({
  subject: s.subject, preheader: s.preheader, editorMode: s.editorMode, bodyHtml: s.bodyHtml,
  customText: s.customText, customTextEnabled: s.customTextEnabled, theme: s.theme, includeSignature: s.includeSignature,
});

export function toComposeContent(c: ContentSlice): ComposeContent {
  return {
    subject: c.subject,
    preheader: c.preheader,
    editorMode: c.editorMode,
    bodyHtml: c.bodyHtml,
    bodyText: c.customTextEnabled && c.customText.trim() ? c.customText : null,
    theme: c.theme,
    includeSignature: c.includeSignature,
  };
}

export function useComposeContent(): ComposeContent {
  const c = useComposeStore(useShallow(pickContent));
  return useMemo(() => toComposeContent(c), [c]);
}

export function useSaveTemplateContent(): Omit<SaveTemplateRequest, 'name' | 'category'> {
  const content = useComposeContent();
  const doc = useComposeStore(s => s.doc);
  const sender = useComposeStore(s => s.senderAccountId);
  return useMemo(() => ({
    subject: content.subject,
    preheader: content.preheader,
    editorMode: content.editorMode,
    docJson: content.editorMode === 'rich' ? doc : null,
    html: content.bodyHtml,
    text: content.bodyText ?? null,
    theme: content.theme,
    includeSignature: content.includeSignature,
    defaultSenderAccountId: sender,
  }), [content, doc, sender]);
}

/** Render request for the preview, or null while no sender is chosen. */
export function useRenderRequest(row: RecipientRow | undefined, unsubscribe?: boolean): RenderRequest | null {
  const content = useComposeContent();
  const senderAccountId = useComposeStore(s => s.senderAccountId);
  return useMemo(() => {
    if (!senderAccountId) return null;
    return { ...content, senderAccountId, row, ...(unsubscribe !== undefined ? { unsubscribe } : {}) };
  }, [content, senderAccountId, row, unsubscribe]);
}

export function useSelectedSender(): { sender: SenderAccount | null; senders: SenderAccount[]; loading: boolean; error: unknown } {
  const q = useMySenders();
  const id = useComposeStore(s => s.senderAccountId);
  const senders = useMemo(() => (q.data ?? []).filter(s => !s.disabled), [q.data]);
  return { sender: senders.find(s => s.id === id) ?? null, senders, loading: q.isLoading, error: q.error };
}

/** Whether "auto" unsubscribe would switch on: more than 50 recipients outside kiit.ac.in. */
export function autoUnsubscribeOn(build: BuildJobRecipientsResult): boolean {
  const external = build.recipients.filter(r => !/@kiit\.ac\.in$/i.test(r.to.split(',')[0].trim())).length;
  return external > 50;
}
