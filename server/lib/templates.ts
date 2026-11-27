// Pure template helpers: the access rule, request validation and import naming.
// No database or Node imports here, so unit tests can load this file without a DB.
import { z } from 'zod';
import type { Role, TemplateVisibility, ThemeId } from '../../shared/types.js';

export const THEME_IDS = ['heritage', 'meridian', 'letterhead', 'slate', 'plain'] as const satisfies readonly ThemeId[];
export const MAX_NAME = 120;
export const MAX_HTML_BYTES = 500 * 1024;
export const MAX_IMPORT = 100;

// ------------------------------------------------------------------ access rule

export interface TemplateAccess {
  /** Appears in the caller's default listing (GET /templates). */
  listed: boolean;
  /** May open it (GET /:id) and duplicate it. Admins may open any template. */
  view: boolean;
  /** May PUT the content. Owner only, admins included. */
  edit: boolean;
  /** May DELETE. Owner or admin. */
  delete: boolean;
  /** May change visibility / share list. Admin only (any template). */
  share: boolean;
  /** Response includes `sharedUserIds`. Owner or admin. */
  seeShares: boolean;
}

/**
 * The single source of truth for who may do what with a template.
 * A user sees their own templates, everything `shared_all`, and `shared_users`
 * templates whose share list contains them. Admins list the same set (plus
 * everything via `?scope=all`), may open, delete and share any template, but
 * can edit only their own.
 */
export function templateAccess(
  t: { ownerId: string; visibility: TemplateVisibility },
  viewer: { id: string; role: Role },
  sharedWithViewer: boolean,
): TemplateAccess {
  const owner = t.ownerId === viewer.id;
  const admin = viewer.role === 'admin';
  const listed = owner || t.visibility === 'shared_all' || (t.visibility === 'shared_users' && sharedWithViewer);
  return {
    listed,
    view: listed || admin,
    edit: owner,
    delete: owner || admin,
    share: admin,
    seeShares: owner || admin,
  };
}

// ------------------------------------------------------------------ validation

const byteLength = (s: string) => new TextEncoder().encode(s).length;
const GUID = z.guid({ error: 'Not a valid id.' });

export const saveTemplateSchema = z.object({
  name: z.string({ error: 'Give the template a name.' }).trim()
    .min(1, 'Give the template a name.')
    .max(MAX_NAME, `Template names can be at most ${MAX_NAME} characters.`),
  category: z.string().trim().max(60, 'Categories can be at most 60 characters.').nullable().optional()
    .transform(v => (v ? v : null)),
  subject: z.string({ error: 'Subject must be text.' }).max(998, 'The subject line is too long (max 998 characters).'),
  preheader: z.string().max(500, 'The preheader is too long (max 500 characters).').optional().default(''),
  editorMode: z.enum(['rich', 'html'], { error: "Editor mode must be 'rich' or 'html'." }),
  docJson: z.unknown().optional().transform(v => (v === undefined ? null : v)),
  html: z.string({ error: 'The template body (html) must be text.' })
    .refine(s => byteLength(s) <= MAX_HTML_BYTES, 'The template body is larger than 500 KB. Remove embedded images or unused markup.'),
  text: z.string().max(200_000, 'The plain-text part is too long.').nullable().optional().transform(v => v ?? null),
  theme: z.enum(THEME_IDS, { error: `Theme must be one of: ${THEME_IDS.join(', ')}.` }),
  includeSignature: z.boolean().optional().default(true),
  defaultSenderAccountId: GUID.nullable().optional().transform(v => v ?? null),
}).refine(v => v.docJson === null || byteLength(JSON.stringify(v.docJson) ?? '') <= 2 * MAX_HTML_BYTES, {
  message: 'The editor document is too large.', path: ['docJson'],
});

export type ValidSaveTemplate = z.output<typeof saveTemplateSchema>;

export const shareSchema = z.object({
  visibility: z.enum(['private', 'shared_all', 'shared_users'], { error: 'Visibility must be private, shared_all or shared_users.' }),
  userIds: z.array(GUID).max(500).optional().default([]),
}).refine(v => v.visibility !== 'shared_users' || v.userIds.length > 0, {
  message: 'Choose at least one person to share with.', path: ['userIds'],
});

export const importSchema = z.object({
  templates: z.array(z.unknown(), { error: 'Expected { templates: [...] }.' })
    .min(1, 'There are no templates to import.')
    .max(MAX_IMPORT, `Import at most ${MAX_IMPORT} templates at a time.`),
});

/** Validate one template from an import batch; error messages name the template. */
export function parseImportItem(raw: unknown, index: number):
  { ok: true; value: ValidSaveTemplate } | { ok: false; error: string } {
  const r = saveTemplateSchema.safeParse(raw);
  if (r.success) return { ok: true, value: r.data };
  const name = raw && typeof raw === 'object' && typeof (raw as { name?: unknown }).name === 'string'
    ? ` ("${(raw as { name: string }).name.slice(0, 60)}")` : '';
  return { ok: false, error: `Template ${index + 1}${name}: ${r.error.issues[0]?.message || 'invalid.'}` };
}

// ------------------------------------------------------------------ naming

const norm = (s: string) => s.trim().toLowerCase();
const clip = (base: string, suffix: string) => base.slice(0, Math.max(1, MAX_NAME - suffix.length)).trimEnd() + suffix;

/** "<name> (copy)", trimmed to fit the 120-character limit. */
export function copyName(name: string): string {
  return clip(name, ' (copy)');
}

/**
 * Names for an import batch. A name already used by the same owner (or earlier
 * in the batch) gets " (imported)", then " (imported 2)", " (imported 3)", ...
 * Comparison is case-insensitive.
 */
export function resolveImportNames(existing: Iterable<string>, incoming: string[]): string[] {
  const taken = new Set([...existing].map(norm));
  return incoming.map((raw) => {
    const name = raw.trim();
    let candidate = name;
    for (let n = 1; taken.has(norm(candidate)); n++) {
      candidate = clip(name, n === 1 ? ' (imported)' : ` (imported ${n})`);
    }
    taken.add(norm(candidate));
    return candidate;
  });
}
