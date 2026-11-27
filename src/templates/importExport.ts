// Pure import/export helpers for the template library (no React, no DOM APIs),
// so they can be unit-tested directly.
//
// Export format (v2):  { version: 2, exportedAt, templates: SaveTemplateRequest[] }
// Legacy v1 format:    { [name]: { subject, preheader, editorMode: 'html'|'rich'|'text', html, text, richHtml } }
//   (both the v1 localStorage key `kiit_mailer_templates` and its "Export all" file).
import type { SaveTemplateRequest, Template, ThemeId } from '../../shared/types';
import { textToHtml } from '../../shared/text';

export const LEGACY_STORAGE_KEY = 'kiit_mailer_templates';
const THEME_IDS: ThemeId[] = ['heritage', 'meridian', 'letterhead', 'slate', 'plain'];

export interface ExportFileV2 {
  version: 2;
  exportedAt?: string;
  templates: SaveTemplateRequest[];
}

export interface LegacyTemplate {
  subject?: string; preheader?: string; editorMode?: 'html' | 'rich' | 'text' | string;
  html?: string; text?: string; richHtml?: string;
}

const str = (v: unknown): string => (typeof v === 'string' ? v : '');
const isObj = (v: unknown): v is Record<string, unknown> => !!v && typeof v === 'object' && !Array.isArray(v);

/** Map one v1 template to a v2 save request. Imports use the 'plain' theme: v1 bodies carried their own styling. */
export function legacyToRequest(name: string, t: LegacyTemplate): SaveTemplateRequest {
  const base = {
    name: name.trim().slice(0, 120) || 'Untitled template',
    category: 'Imported (v1)',
    subject: str(t.subject),
    preheader: str(t.preheader),
    theme: 'plain' as ThemeId,
    includeSignature: false,
    defaultSenderAccountId: null,
  };
  switch (t.editorMode) {
    case 'rich':
      return { ...base, editorMode: 'rich', docJson: null, html: str(t.richHtml), text: null };
    case 'text': {
      const text = str(t.text);
      return { ...base, editorMode: 'rich', docJson: null, html: `<p>${textToHtml(text)}</p>`, text };
    }
    case 'html':
    default:
      // v1 defaulted to html mode when editorMode was missing.
      return { ...base, editorMode: 'html', docJson: null, html: str(t.html), text: null };
  }
}

/** Map a whole v1 object (localStorage value or "Export all" file). */
export function legacyToRequests(obj: Record<string, unknown>): SaveTemplateRequest[] {
  return Object.entries(obj)
    .filter(([, v]) => isObj(v))
    .map(([name, v]) => legacyToRequest(name, v as LegacyTemplate));
}

/** Snapshot of a full template as a save request (used by export, rename and category edits). */
export function templateToRequest(t: Template): SaveTemplateRequest {
  return {
    name: t.name, category: t.category, subject: t.subject, preheader: t.preheader,
    editorMode: t.editorMode, docJson: t.docJson ?? null, html: t.html, text: t.text,
    theme: t.theme, includeSignature: t.includeSignature, defaultSenderAccountId: t.defaultSenderAccountId,
  };
}

export function buildExport(templates: Template[], now = new Date()): ExportFileV2 {
  return { version: 2, exportedAt: now.toISOString(), templates: templates.map(templateToRequest) };
}

function normaliseV2(raw: unknown, i: number): SaveTemplateRequest {
  if (!isObj(raw)) throw new Error(`Template ${i + 1} is not an object.`);
  const name = str(raw.name).trim();
  if (!name) throw new Error(`Template ${i + 1} has no name.`);
  const editorMode = raw.editorMode === 'html' ? 'html' : 'rich';
  const theme = THEME_IDS.includes(raw.theme as ThemeId) ? (raw.theme as ThemeId) : 'plain';
  return {
    name,
    category: typeof raw.category === 'string' && raw.category.trim() ? raw.category.trim() : null,
    subject: str(raw.subject),
    preheader: str(raw.preheader),
    editorMode,
    docJson: raw.docJson ?? null,
    html: str(raw.html),
    text: typeof raw.text === 'string' ? raw.text : null,
    theme,
    includeSignature: typeof raw.includeSignature === 'boolean' ? raw.includeSignature : true,
    defaultSenderAccountId: typeof raw.defaultSenderAccountId === 'string' ? raw.defaultSenderAccountId : null,
  };
}

export type ImportFormat = 'v2' | 'legacy';

/**
 * Parse an import file's JSON (already JSON.parse'd). Accepts the v2 export format
 * and the legacy v1 object. Throws an Error with a plain-English message otherwise.
 */
export function parseImport(data: unknown): { format: ImportFormat; templates: SaveTemplateRequest[] } {
  if (isObj(data) && data.version === 2 && Array.isArray(data.templates)) {
    return { format: 'v2', templates: data.templates.map(normaliseV2) };
  }
  if (Array.isArray(data)) {
    // Tolerate a bare array of v2 templates.
    return { format: 'v2', templates: data.map(normaliseV2) };
  }
  if (isObj(data)) {
    const entries = Object.values(data);
    const looksLegacy = entries.length > 0 && entries.every(v => isObj(v) && ('subject' in v || 'html' in v || 'richHtml' in v || 'text' in v || 'editorMode' in v));
    if (looksLegacy) return { format: 'legacy', templates: legacyToRequests(data) };
    if (entries.length === 0) return { format: 'legacy', templates: [] };
  }
  throw new Error('This file is not a KIIT Mailer template export. Use a file made by "Export" here, or by "Export all" in KIIT Mailer v1.');
}

/** Parse a file's text. */
export function parseImportText(text: string) {
  let data: unknown;
  try { data = JSON.parse(text); } catch { throw new Error('This file is not valid JSON.'); }
  return parseImport(data);
}

/** Client-side mirror of the server's duplicate naming: "<name> (copy)" within 120 characters. */
export function copyNameFor(name: string): string {
  const suffix = ' (copy)';
  return name.slice(0, 120 - suffix.length).trimEnd() + suffix;
}

/** Split into request-sized chunks (the API accepts at most 100 per import). */
export function chunk<T>(items: T[], size = 100): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < items.length; i += size) out.push(items.slice(i, i + size));
  return out;
}
