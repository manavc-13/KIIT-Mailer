// /api/admin/audit: the append-only audit trail, filtered and cursor-paginated
// (newest first by `at`, then id), plus a CSV export with the same filters.
import { Hono } from 'hono';
import { and, desc, eq, gte, like, lt, sql, type SQL } from 'drizzle-orm';
import { z } from 'zod';
import { db, schema } from '../db/client.js';
import type { AppEnv } from '../types.js';
import { badRequest, query } from '../lib/http.js';
import { toCSV } from '../../shared/text.js';
import type { AuditEntry } from '../../shared/types.js';

const PAGE_SIZE = 50;
const EXPORT_MAX = 20_000;
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const FilterSchema = z.object({
  actorId: z.string().regex(UUID_RE, 'Unknown user.').optional().or(z.literal('').transform(() => undefined)),
  /** Exact action ("user.create") or a category prefix ending in ".*" ("user.*"). */
  action: z.string().max(64).regex(/^[a-z_]+(\.[a-z_*]+)?$/, 'Unknown action.').optional().or(z.literal('').transform(() => undefined)),
  /** ISO timestamp or YYYY-MM-DD (inclusive). */
  from: z.string().max(40).optional(),
  /** ISO timestamp (exclusive) or YYYY-MM-DD (inclusive: the whole day is included). */
  to: z.string().max(40).optional(),
  cursor: z.string().max(200).optional(),
});
type Filters = z.infer<typeof FilterSchema>;

const DATE_ONLY = /^\d{4}-\d{2}-\d{2}$/;
function parseDate(v: string | undefined, label: string, endOfDay: boolean): Date | null {
  if (!v) return null;
  const d = new Date(v);
  if (Number.isNaN(d.getTime())) throw badRequest(`The "${label}" date is not valid.`);
  if (endOfDay && DATE_ONLY.test(v)) d.setUTCDate(d.getUTCDate() + 1);
  return d;
}

function encodeCursor(atText: string, id: string): string {
  return Buffer.from(`${atText}|${id}`, 'utf8').toString('base64url');
}
function decodeCursor(cursor: string): { at: string; id: string } {
  const [at, id] = Buffer.from(cursor, 'base64url').toString('utf8').split('|');
  if (!at || !id || !UUID_RE.test(id) || Number.isNaN(new Date(at).getTime())) throw badRequest('Invalid page cursor.');
  return { at, id };
}

function whereFor(f: Filters, withCursor: boolean): SQL | undefined {
  const a = schema.auditLog;
  const conds: (SQL | undefined)[] = [];
  if (f.actorId) conds.push(eq(a.actorId, f.actorId));
  if (f.action) conds.push(f.action.endsWith('.*') ? like(a.action, f.action.slice(0, -1) + '%') : eq(a.action, f.action));
  const from = parseDate(f.from, 'from', false);
  const to = parseDate(f.to, 'to', true);
  if (from) conds.push(gte(a.at, from));
  if (to) conds.push(lt(a.at, to));
  if (withCursor && f.cursor) {
    const cur = decodeCursor(f.cursor);
    // Compare at full microsecond precision using the text form handed out in the cursor.
    conds.push(sql`(${a.at}, ${a.id}) < (${cur.at}::timestamptz, ${cur.id}::uuid)`);
  }
  return conds.length ? and(...conds) : undefined;
}

const selection = {
  row: schema.auditLog,
  atText: sql<string>`${schema.auditLog.at}::text`.as('at_text'),
};

function toEntry(r: typeof schema.auditLog.$inferSelect): AuditEntry {
  return {
    id: r.id, actorId: r.actorId, actorEmail: r.actorEmail, action: r.action,
    targetType: r.targetType, targetId: r.targetId,
    meta: (r.meta as Record<string, unknown> | null) ?? null, ip: r.ip, at: r.at.toISOString(),
  };
}

const r = new Hono<AppEnv>();

r.get('/', async (c) => {
  const f = query(c, FilterSchema);
  const rows = await db.select(selection).from(schema.auditLog)
    .where(whereFor(f, true))
    .orderBy(desc(schema.auditLog.at), desc(schema.auditLog.id))
    .limit(PAGE_SIZE + 1);
  const page = rows.slice(0, PAGE_SIZE);
  const last = page[page.length - 1];
  const nextCursor = rows.length > PAGE_SIZE && last ? encodeCursor(last.atText, last.row.id) : null;
  return c.json<{ entries: AuditEntry[]; nextCursor: string | null }>({ entries: page.map(x => toEntry(x.row)), nextCursor });
});

/** Neutralise spreadsheet formula injection (cells starting with = + - @). */
const safeCell = (v: unknown) => {
  const s = String(v ?? '');
  return /^[=+\-@\t\r]/.test(s) ? `'${s}` : s;
};

r.get('/export.csv', async (c) => {
  const f = query(c, FilterSchema);
  const rows = await db.select().from(schema.auditLog)
    .where(whereFor(f, false))
    .orderBy(desc(schema.auditLog.at), desc(schema.auditLog.id))
    .limit(EXPORT_MAX);
  const columns = ['at', 'actor_email', 'actor_id', 'action', 'target_type', 'target_id', 'ip', 'meta'];
  const csv = toCSV(columns, rows.map(x => ({
    at: x.at.toISOString(),
    actor_email: safeCell(x.actorEmail),
    actor_id: x.actorId ?? '',
    action: safeCell(x.action),
    target_type: safeCell(x.targetType),
    target_id: safeCell(x.targetId),
    ip: safeCell(x.ip),
    meta: x.meta == null ? '' : safeCell(JSON.stringify(x.meta)),
  })));
  const stamp = new Date().toISOString().slice(0, 10);
  return c.body(csv, 200, {
    'Content-Type': 'text/csv; charset=utf-8',
    'Content-Disposition': `attachment; filename="kiit-mailer-audit-${stamp}.csv"`,
  });
});

export default r;
