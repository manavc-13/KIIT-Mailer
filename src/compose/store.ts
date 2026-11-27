// Compose flow state (zustand). One module-level store, so leaving Compose and coming back keeps
// the work in progress; the draft is also autosaved per user to localStorage (see draft.ts).
import { create } from 'zustand';
import type { EditorMode, JobMode, SenderAccount, Template, ThemeId, UploadRef } from '../../shared/types';
import { DEFAULT_THEME } from '../../shared/themes';
import {
  DEFAULT_GRID_COLUMNS, DEFAULT_GRID_ROLE_MAP, EMPTY_ROLE_MAP, detectRoleMap, sanitizeRoleMap,
  type RecipientRow, type RoleMap,
} from '../../shared/recipients';

export type RecipientSource = 'single' | 'csv' | 'manual';
export type ComposeStep = 'recipients' | 'compose' | 'review';
export const STEPS: { id: ComposeStep; label: string }[] = [
  { id: 'recipients', label: 'Recipients' },
  { id: 'compose', label: 'Compose' },
  { id: 'review', label: 'Review & send' },
];

export interface ExtraField { id: string; key: string; value: string }

export interface AttachmentItem {
  localId: string;
  filename: string;
  size: number;
  contentType: string;
  status: 'uploading' | 'ready' | 'error';
  /** 0..1 upload progress. */
  progress: number;
  ref: UploadRef | null;
  error: string | null;
}

export interface LoadedTemplateMeta { id: string; name: string; category: string | null; isOwner: boolean }
export type UnsubscribeOption = 'auto' | true | false;

export interface ComposeData {
  step: ComposeStep;
  // recipients
  source: RecipientSource;
  singleName: string;
  singleEmail: string;
  singleExtras: ExtraField[];
  csvFileName: string | null;
  csvRows: RecipientRow[] | null;
  csvHeaders: string[];
  csvRoleMap: RoleMap;
  /** A draft was restored without its (large) CSV rows. */
  csvNeedsReupload: boolean;
  gridColumns: string[];
  gridRows: RecipientRow[];
  gridRoleMap: RoleMap;
  globalCc: string;
  globalBcc: string;
  mode: JobMode;
  dedupe: boolean;
  // sender + template
  senderAccountId: string | null;
  templateId: string | null;
  templateMeta: LoadedTemplateMeta | null;
  // content
  subject: string;
  preheader: string;
  editorMode: EditorMode;
  doc: unknown | null;
  bodyHtml: string;
  /** Seed for the rich editor when only HTML exists (template without doc, or html->rich switch). */
  seedHtml: string | null;
  customTextEnabled: boolean;
  customText: string;
  theme: ThemeId;
  /** The user picked a theme explicitly (sender switches no longer override it). */
  themeTouched: boolean;
  includeSignature: boolean;
  // sending options
  attachments: AttachmentItem[];
  unsubscribe: UnsubscribeOption;
}

interface ComposeSession {
  /** User id whose work is in the store (a different sign-in resets it). */
  ownerId: string | null;
  /** Bumped to remount the editor with new initial content. */
  editorKey: number;
  /** Content changed since the last template load / save. */
  dirty: boolean;
  previewIndex: number;
  /** A test email went out successfully in this session (Review checklist). */
  testSent: boolean;
}

export type ComposeState = ComposeData & ComposeSession & ComposeActions;

interface ComposeActions {
  set: (patch: Partial<ComposeData & ComposeSession>) => void;
  /** Content edits: also mark the compose dirty. */
  edit: (patch: Partial<ComposeData>) => void;
  reset: (ownerId?: string | null) => void;
  setStep: (step: ComposeStep) => void;
  // single
  addExtra: () => void;
  updateExtra: (id: string, patch: Partial<Omit<ExtraField, 'id'>>) => void;
  removeExtra: (id: string) => void;
  // csv
  loadCsv: (fileName: string, headers: string[], rows: RecipientRow[]) => void;
  clearCsv: () => void;
  setCsvRole: (role: keyof RoleMap, column: string | null) => void;
  // grid
  setGridCell: (rowIndex: number, column: string, value: string) => void;
  addGridRow: () => void;
  removeGridRows: (indices: number[]) => void;
  addGridColumn: (name: string) => boolean;
  renameGridColumn: (index: number, name: string) => boolean;
  removeGridColumn: (index: number) => void;
  setGridRole: (role: keyof RoleMap, column: string | null) => void;
  importGrid: (columns: string[], rows: RecipientRow[], append: boolean) => void;
  pasteIntoGrid: (startRow: number, startColumn: string, text: string) => number;
  resetGrid: () => void;
  // sender + template
  selectSender: (sender: SenderAccount) => void;
  applyTemplate: (t: Template, mySenderIds: string[]) => void;
  setTemplateMeta: (t: Template) => void;
  // attachments
  addAttachment: (item: AttachmentItem) => void;
  updateAttachment: (localId: string, patch: Partial<AttachmentItem>) => void;
  removeAttachment: (localId: string) => void;
}

let idCounter = 0;
export const newId = (prefix = 'id') => `${prefix}_${Date.now().toString(36)}_${(idCounter++).toString(36)}`;

const blankRow = (columns: string[]): RecipientRow => Object.fromEntries(columns.map(c => [c, '']));

export function initialData(): ComposeData {
  return {
    step: 'recipients',
    source: 'single',
    singleName: '',
    singleEmail: '',
    singleExtras: [],
    csvFileName: null,
    csvRows: null,
    csvHeaders: [],
    csvRoleMap: { ...EMPTY_ROLE_MAP },
    csvNeedsReupload: false,
    gridColumns: [...DEFAULT_GRID_COLUMNS],
    gridRows: [blankRow(DEFAULT_GRID_COLUMNS)],
    gridRoleMap: { ...DEFAULT_GRID_ROLE_MAP },
    globalCc: '',
    globalBcc: '',
    mode: 'personal',
    dedupe: true,
    senderAccountId: null,
    templateId: null,
    templateMeta: null,
    subject: '',
    preheader: '',
    editorMode: 'rich',
    doc: null,
    bodyHtml: '',
    seedHtml: null,
    customTextEnabled: false,
    customText: '',
    theme: DEFAULT_THEME,
    themeTouched: false,
    includeSignature: true,
    attachments: [],
    unsubscribe: 'auto',
  };
}

const initialSession = (ownerId: string | null = null): ComposeSession => ({
  ownerId, editorKey: 0, dirty: false, previewIndex: 0, testSent: false,
});

export const useComposeStore = create<ComposeState>()((set, get) => ({
  ...initialData(),
  ...initialSession(),

  set: patch => set(patch),
  edit: patch => set({ ...patch, dirty: true }),
  reset: ownerId => set(s => ({
    ...initialData(), ...initialSession(ownerId === undefined ? s.ownerId : ownerId), editorKey: s.editorKey + 1,
  })),
  setStep: step => set({ step }),

  addExtra: () => set(s => ({ singleExtras: [...s.singleExtras, { id: newId('f'), key: '', value: '' }] })),
  updateExtra: (id, patch) => set(s => ({ singleExtras: s.singleExtras.map(f => (f.id === id ? { ...f, ...patch } : f)) })),
  removeExtra: id => set(s => ({ singleExtras: s.singleExtras.filter(f => f.id !== id) })),

  loadCsv: (fileName, headers, rows) => set({
    csvFileName: fileName, csvHeaders: headers, csvRows: rows, csvRoleMap: detectRoleMap(headers),
    csvNeedsReupload: false, previewIndex: 0,
  }),
  clearCsv: () => set({ csvFileName: null, csvHeaders: [], csvRows: null, csvRoleMap: { ...EMPTY_ROLE_MAP }, csvNeedsReupload: false }),
  setCsvRole: (role, column) => set(s => ({ csvRoleMap: { ...s.csvRoleMap, [role]: column } })),

  setGridCell: (ri, col, value) => set(s => {
    const rows = s.gridRows.slice();
    rows[ri] = { ...rows[ri], [col]: value };
    return { gridRows: rows };
  }),
  addGridRow: () => set(s => ({ gridRows: [...s.gridRows, blankRow(s.gridColumns)] })),
  removeGridRows: indices => set(s => {
    const drop = new Set(indices);
    const rows = s.gridRows.filter((_, i) => !drop.has(i));
    return { gridRows: rows.length ? rows : [blankRow(s.gridColumns)] };
  }),
  addGridColumn: name => {
    const clean = name.trim();
    const s = get();
    if (!clean || s.gridColumns.some(c => c.toLowerCase() === clean.toLowerCase())) return false;
    set({ gridColumns: [...s.gridColumns, clean], gridRows: s.gridRows.map(r => ({ ...r, [clean]: '' })) });
    return true;
  },
  renameGridColumn: (index, name) => {
    const s = get();
    const clean = name.trim();
    const old = s.gridColumns[index];
    if (!clean || old === undefined) return false;
    if (clean === old) return true;
    if (s.gridColumns.some((c, i) => i !== index && c.toLowerCase() === clean.toLowerCase())) return false;
    const cols = s.gridColumns.slice(); cols[index] = clean;
    const rows = s.gridRows.map(r => { const { [old]: v, ...rest } = r; return { ...rest, [clean]: v ?? '' }; });
    const rm = { ...s.gridRoleMap };
    (Object.keys(rm) as (keyof RoleMap)[]).forEach(k => { if (rm[k] === old) rm[k] = clean; });
    set({ gridColumns: cols, gridRows: rows, gridRoleMap: rm });
    return true;
  },
  removeGridColumn: index => set(s => {
    const col = s.gridColumns[index];
    if (col === undefined) return {};
    const cols = s.gridColumns.filter((_, i) => i !== index);
    const rows = s.gridRows.map(r => { const { [col]: _drop, ...rest } = r; return rest; });
    return { gridColumns: cols, gridRows: rows, gridRoleMap: sanitizeRoleMap(s.gridRoleMap, cols) };
  }),
  setGridRole: (role, column) => set(s => ({ gridRoleMap: { ...s.gridRoleMap, [role]: column } })),
  importGrid: (columns, rows, append) => set(s => {
    let cols: string[];
    let base: RecipientRow[];
    if (append) {
      cols = [...s.gridColumns];
      columns.forEach(c => { if (!cols.includes(c)) cols.push(c); });
      base = s.gridRows.filter(r => Object.values(r).some(v => String(v ?? '').trim()));
    } else {
      cols = columns.filter(Boolean);
      base = [];
    }
    const norm = (r: RecipientRow) => Object.fromEntries(cols.map(c => [c, r[c] ?? '']));
    const all = [...base.map(norm), ...rows.map(norm)];
    const detected = detectRoleMap(cols);
    // Fall back to an Email column so the grid always has somewhere to type an address.
    if (!detected.email) { cols.push('Email'); all.forEach(r => { r.Email = ''; }); detected.email = 'Email'; }
    return { gridColumns: cols, gridRows: all.length ? all : [blankRow(cols)], gridRoleMap: detected };
  }),
  pasteIntoGrid: (startRow, startColumn, text) => {
    const s = get();
    const startCi = s.gridColumns.indexOf(startColumn);
    if (startCi < 0) return 0;
    const lines = text.replace(/\r/g, '').replace(/\n+$/, '').split('\n');
    const rows = s.gridRows.map(r => ({ ...r }));
    lines.forEach((line, li) => {
      const target = startRow + li;
      while (rows.length <= target) rows.push(blankRow(s.gridColumns));
      line.split('\t').forEach((val, k) => {
        const col = s.gridColumns[startCi + k];
        if (col !== undefined) rows[target][col] = val.trim();
      });
    });
    set({ gridRows: rows });
    return lines.length;
  },
  resetGrid: () => set({
    gridColumns: [...DEFAULT_GRID_COLUMNS], gridRows: [blankRow(DEFAULT_GRID_COLUMNS)], gridRoleMap: { ...DEFAULT_GRID_ROLE_MAP },
  }),

  selectSender: sender => set(s => ({
    senderAccountId: sender.id,
    // The theme follows the sender's letterhead unless a template is loaded or the user chose one.
    ...(!s.templateId && !s.themeTouched ? { theme: sender.defaultTheme } : {}),
  })),
  applyTemplate: (t, mySenderIds) => set(s => ({
    templateId: t.id,
    templateMeta: { id: t.id, name: t.name, category: t.category, isOwner: t.isOwner },
    subject: t.subject ?? '',
    preheader: t.preheader ?? '',
    editorMode: t.editorMode,
    doc: t.editorMode === 'rich' ? (t.docJson ?? null) : null,
    seedHtml: t.editorMode === 'rich' && !t.docJson ? t.html : null,
    bodyHtml: t.html ?? '',
    customText: t.text ?? '',
    customTextEnabled: !!(t.text && t.text.trim()),
    theme: t.theme,
    themeTouched: true,
    includeSignature: t.includeSignature ?? true,
    senderAccountId: t.defaultSenderAccountId && mySenderIds.includes(t.defaultSenderAccountId)
      ? t.defaultSenderAccountId : s.senderAccountId,
    editorKey: s.editorKey + 1,
    dirty: false,
    testSent: false,
  })),
  setTemplateMeta: t => set({
    templateId: t.id, templateMeta: { id: t.id, name: t.name, category: t.category, isOwner: t.isOwner }, dirty: false,
  }),

  addAttachment: item => set(s => ({ attachments: [...s.attachments, item] })),
  updateAttachment: (localId, patch) => set(s => ({
    attachments: s.attachments.map(a => (a.localId === localId ? { ...a, ...patch } : a)),
  })),
  removeAttachment: localId => set(s => ({ attachments: s.attachments.filter(a => a.localId !== localId) })),
}));

/** True when the compose holds content worth protecting before replacing it. */
export function hasComposeContent(s: Pick<ComposeData, 'subject' | 'bodyHtml' | 'preheader' | 'customText'>): boolean {
  const body = s.bodyHtml.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
  return !!(s.subject.trim() || body || s.preheader.trim() || s.customText.trim());
}
