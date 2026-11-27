// /api/templates: personal and shared email templates.
// Access is decided by templateAccess() in server/lib/templates.ts (pure, unit-tested);
// the SQL here only pre-filters to a superset and the rule makes the final call.
import { Hono } from 'hono';
import { desc, eq, inArray, ne, or } from 'drizzle-orm';
import { z } from 'zod';
import { db, schema } from '../db/client.js';
import type { AppEnv, UserRow } from '../types.js';
import { me, requireUser } from '../lib/session.js';
import { ApiError, badRequest, body, forbidden, iso, notFound, query } from '../lib/http.js';
import { audit } from '../lib/audit.js';
import { assertSenderAccess } from '../lib/access.js';
import {
  copyName, importSchema, parseImportItem, resolveImportNames, saveTemplateSchema, shareSchema, templateAccess,
  type ValidSaveTemplate,
} from '../lib/templates.js';
import type { Template, TemplateSummary, ThemeId } from '../../shared/types.js';

const r = new Hono<AppEnv>();
r.use('*', requireUser);

const T = schema.templates;
type TemplateRow = typeof T.$inferSelect;

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const summaryCols = {
  id: T.id, name: T.name, category: T.category, subject: T.subject, theme: T.theme, editorMode: T.editorMode,
  visibility: T.visibility, ownerId: T.ownerId, defaultSenderAccountId: T.defaultSenderAccountId, updatedAt: T.updatedAt,
  ownerName: schema.users.name,
};

function toSummary(row: Pick<TemplateRow, 'id' | 'name' | 'category' | 'subject' | 'theme' | 'editorMode' | 'visibility' | 'ownerId' | 'defaultSenderAccountId' | 'updatedAt'> & { ownerName: string }, viewer: UserRow): TemplateSummary {
  return {
    id: row.id, name: row.name, category: row.category, subject: row.subject, theme: row.theme as ThemeId,
    editorMode: row.editorMode, visibility: row.visibility, ownerId: row.ownerId, ownerName: row.ownerName,
    isOwner: row.ownerId === viewer.id, defaultSenderAccountId: row.defaultSenderAccountId,
    updatedAt: iso(row.updatedAt)!,
  };
}

async function shareIdsOf(templateId: string): Promise<string[]> {
  const rows = await db.select({ u: schema.templateShares.userId }).from(schema.templateShares)
    .where(eq(schema.templateShares.templateId, templateId));
  return rows.map(x => x.u);
}

/** Loads a template with its owner name and share list, and the caller's access to it. 404 if not viewable. */
async function load(id: string, viewer: UserRow) {
  if (!UUID_RE.test(id)) throw notFound('Template not found.');
  const [row] = await db.select({ t: T, ownerName: schema.users.name }).from(T)
    .innerJoin(schema.users, eq(schema.users.id, T.ownerId))
    .where(eq(T.id, id)).limit(1);
  if (!row) throw notFound('Template not found.');
  const shares = await shareIdsOf(id);
  const access = templateAccess(row.t, viewer, shares.includes(viewer.id));
  // Hide existence from people who may not see it.
  if (!access.view) throw notFound('Template not found.');
  return { row: row.t, ownerName: row.ownerName, shares, access };
}

function toTemplate(row: TemplateRow, ownerName: string, viewer: UserRow, shares: string[]): Template {
  const access = templateAccess(row, viewer, shares.includes(viewer.id));
  return {
    ...toSummary({ ...row, ownerName }, viewer),
    preheader: row.preheader, docJson: row.docJson ?? null, html: row.html, text: row.text,
    includeSignature: row.includeSignature,
    sharedUserIds: access.seeShares ? shares : [],
  };
}

/** A default sender must be an account the caller may send from; otherwise a clear 400. */
async function checkDefaultSender(user: UserRow, senderId: string | null): Promise<void> {
  if (!senderId) return;
  try {
    await assertSenderAccess(user, senderId);
  } catch (err) {
    const why = err instanceof ApiError ? err.message : 'It could not be checked.';
    throw new ApiError(400, `The default sender for this template is not an account you can send from. ${why}`, 'invalid_sender');
  }
}

/** Default sender kept only when the caller can use it (duplicate/import: never fail the whole action). */
async function usableSender(user: UserRow, senderId: string | null): Promise<string | null> {
  if (!senderId) return null;
  try { await assertSenderAccess(user, senderId); return senderId; } catch { return null; }
}

const values = (v: ValidSaveTemplate) => ({
  name: v.name, category: v.category, subject: v.subject, preheader: v.preheader, editorMode: v.editorMode,
  docJson: v.docJson, html: v.html, text: v.text, theme: v.theme, includeSignature: v.includeSignature,
  defaultSenderAccountId: v.defaultSenderAccountId,
});

// ------------------------------------------------------------------ list

r.get('/', async (c) => {
  const user = me(c);
  const { scope } = query(c, z.object({ scope: z.enum(['mine', 'all']).optional() }));
  if (scope === 'all' && user.role !== 'admin') throw forbidden('Only administrators can list every template.');

  const base = db.select(summaryCols).from(T).innerJoin(schema.users, eq(schema.users.id, T.ownerId));
  if (scope === 'all') {
    const rows = await base.orderBy(desc(T.updatedAt));
    return c.json<{ templates: TemplateSummary[] }>({ templates: rows.map(x => toSummary(x, user)) });
  }
  // Superset: mine or anything not private; templateAccess() then drops shared_users rows not shared with me.
  const [rows, mine] = await Promise.all([
    base.where(or(eq(T.ownerId, user.id), ne(T.visibility, 'private'))).orderBy(desc(T.updatedAt)),
    db.select({ id: schema.templateShares.templateId }).from(schema.templateShares)
      .where(eq(schema.templateShares.userId, user.id)),
  ]);
  const sharedWithMe = new Set(mine.map(x => x.id));
  const visible = rows.filter(x => templateAccess(x, user, sharedWithMe.has(x.id)).listed);
  return c.json<{ templates: TemplateSummary[] }>({ templates: visible.map(x => toSummary(x, user)) });
});

// ------------------------------------------------------------------ import (before /:id)

r.post('/import', async (c) => {
  const user = me(c);
  const { templates: rawItems } = await body(c, importSchema);
  const items: ValidSaveTemplate[] = [];
  rawItems.forEach((raw, i) => {
    const p = parseImportItem(raw, i);
    if (!p.ok) throw badRequest(p.error);
    items.push(p.value);
  });

  const existing = await db.select({ name: T.name }).from(T).where(eq(T.ownerId, user.id));
  const names = resolveImportNames(existing.map(x => x.name), items.map(x => x.name));

  // Resolve default senders once per distinct id; drop those the caller cannot use.
  const senderOk = new Map<string, string | null>();
  for (const id of new Set(items.map(x => x.defaultSenderAccountId).filter((x): x is string => !!x))) {
    senderOk.set(id, await usableSender(user, id));
  }

  const now = new Date();
  const inserted = await db.insert(T).values(items.map((v, i) => ({
    ...values(v),
    name: names[i],
    defaultSenderAccountId: v.defaultSenderAccountId ? senderOk.get(v.defaultSenderAccountId) ?? null : null,
    ownerId: user.id, visibility: 'private' as const, createdAt: now, updatedAt: now,
  }))).returning({ id: T.id });

  const renamed = names.filter((n, i) => n !== items[i].name).length;
  await audit(c, 'template.create', {
    targetType: 'template', meta: { imported: inserted.length, renamed, ids: inserted.map(x => x.id) },
  });
  return c.json<{ created: number; renamed: number }>({ created: inserted.length, renamed });
});

// ------------------------------------------------------------------ CRUD

r.get('/:id', async (c) => {
  const user = me(c);
  const t = await load(c.req.param('id'), user);
  return c.json<{ template: Template }>({ template: toTemplate(t.row, t.ownerName, user, t.shares) });
});

r.post('/', async (c) => {
  const user = me(c);
  const v = await body(c, saveTemplateSchema);
  await checkDefaultSender(user, v.defaultSenderAccountId);
  const [row] = await db.insert(T).values({ ...values(v), ownerId: user.id, visibility: 'private' }).returning();
  await audit(c, 'template.create', { targetType: 'template', targetId: row.id, meta: { name: row.name } });
  return c.json<{ template: Template }>({ template: toTemplate(row, user.name, user, []) }, 201);
});

r.put('/:id', async (c) => {
  const user = me(c);
  const t = await load(c.req.param('id'), user);
  if (!t.access.edit) throw forbidden('Only the owner can edit this template. Duplicate it to make your own copy.');
  const v = await body(c, saveTemplateSchema);
  // Only re-check the default sender when it changes, so a later revocation does not block unrelated edits.
  if (v.defaultSenderAccountId !== t.row.defaultSenderAccountId) await checkDefaultSender(user, v.defaultSenderAccountId);
  const [row] = await db.update(T).set({ ...values(v), updatedAt: new Date() }).where(eq(T.id, t.row.id)).returning();
  await audit(c, 'template.update', {
    targetType: 'template', targetId: row.id,
    meta: { name: row.name, ...(row.name !== t.row.name ? { previousName: t.row.name } : {}) },
  });
  return c.json<{ template: Template }>({ template: toTemplate(row, t.ownerName, user, t.shares) });
});

r.delete('/:id', async (c) => {
  const user = me(c);
  const t = await load(c.req.param('id'), user);
  if (!t.access.delete) throw forbidden('Only the owner or an administrator can delete this template.');
  await db.delete(T).where(eq(T.id, t.row.id));
  await audit(c, 'template.delete', {
    targetType: 'template', targetId: t.row.id, meta: { name: t.row.name, ownerId: t.row.ownerId },
  });
  return c.json({ ok: true as const });
});

r.post('/:id/duplicate', async (c) => {
  const user = me(c);
  const t = await load(c.req.param('id'), user);
  const src = t.row;
  const [row] = await db.insert(T).values({
    ownerId: user.id, visibility: 'private',
    name: copyName(src.name), category: src.category, subject: src.subject, preheader: src.preheader,
    editorMode: src.editorMode, docJson: src.docJson, html: src.html, text: src.text, theme: src.theme,
    includeSignature: src.includeSignature,
    defaultSenderAccountId: await usableSender(user, src.defaultSenderAccountId),
  }).returning();
  await audit(c, 'template.create', {
    targetType: 'template', targetId: row.id, meta: { name: row.name, duplicatedFrom: src.id },
  });
  return c.json<{ template: Template }>({ template: toTemplate(row, user.name, user, []) }, 201);
});

r.put('/:id/share', async (c) => {
  const user = me(c);
  const t = await load(c.req.param('id'), user);
  if (!t.access.share) throw forbidden('Only administrators can share templates.');
  const { visibility, userIds } = await body(c, shareSchema);

  // Owners always see their own template, so they never need a share row.
  let ids = visibility === 'shared_users' ? [...new Set(userIds)].filter(id => id !== t.row.ownerId) : [];
  if (ids.length) {
    const found = await db.select({ id: schema.users.id }).from(schema.users).where(inArray(schema.users.id, ids));
    const known = new Set(found.map(x => x.id));
    const unknown = ids.filter(id => !known.has(id));
    if (unknown.length) throw badRequest(`${unknown.length} of the selected people no longer exist. Refresh and try again.`);
  }
  if (visibility === 'shared_users' && !ids.length) throw badRequest('Choose at least one person other than the owner.');

  await db.transaction(async (tx) => {
    await tx.delete(schema.templateShares).where(eq(schema.templateShares.templateId, t.row.id));
    if (ids.length) await tx.insert(schema.templateShares).values(ids.map(userId => ({ templateId: t.row.id, userId })));
    await tx.update(T).set({ visibility }).where(eq(T.id, t.row.id));
  });
  await audit(c, 'template.share', {
    targetType: 'template', targetId: t.row.id,
    meta: {
      name: t.row.name, visibility, userIds: ids,
      previousVisibility: t.row.visibility, previousUserIds: t.shares,
    },
  });
  ids = await shareIdsOf(t.row.id);
  return c.json<{ template: Template }>({ template: toTemplate({ ...t.row, visibility }, t.ownerName, user, ids) });
});

export default r;
