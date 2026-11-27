// /api/admin/users: create users, set roles, disable/enable, reset passwords and
// allot sender accounts. requireAdmin is applied to /api/admin/* in server/app.ts.
import { Hono } from 'hono';
import { and, asc, eq, inArray, ne } from 'drizzle-orm';
import { z } from 'zod';
import { db, schema } from '../db/client.js';
import type { AppEnv, UserRow } from '../types.js';
import { badRequest, body, conflict, notFound } from '../lib/http.js';
import { me, revokeUserSessions } from '../lib/session.js';
import { audit } from '../lib/audit.js';
import { tempPassword } from '../lib/crypto.js';
import { hashPassword } from '../lib/passwords.js';
import { toPublicUser } from './auth.js';
import { isValidEmail } from '../../shared/text.js';
import type { AdminUser } from '../../shared/types.js';

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const EmailField = z.string().trim().toLowerCase().max(320)
  .refine(isValidEmail, 'Enter a valid email address.');
const NameField = z.string().trim().min(1, 'Enter a name.').max(120, 'Keep the name under 120 characters.');
const RoleField = z.enum(['admin', 'sender']);
const SenderIdsField = z.array(z.string().regex(UUID_RE, 'Unknown sender account.')).max(500);

const CreateSchema = z.object({
  email: EmailField, name: NameField, role: RoleField, senderAccountIds: SenderIdsField.optional(),
});
const UpdateSchema = z.object({
  name: NameField.optional(), role: RoleField.optional(), disabled: z.boolean().optional(),
});
const SendersSchema = z.object({ senderAccountIds: SenderIdsField });

async function grantsFor(userIds: string[]): Promise<Map<string, string[]>> {
  const map = new Map<string, string[]>();
  if (!userIds.length) return map;
  const rows = await db.select().from(schema.userSenderAccess).where(inArray(schema.userSenderAccess.userId, userIds));
  for (const g of rows) map.set(g.userId, [...(map.get(g.userId) ?? []), g.senderAccountId]);
  return map;
}

function toAdminUser(u: UserRow, senderAccountIds: string[]): AdminUser {
  return { ...toPublicUser(u), senderAccountIds };
}

async function loadUser(id: string): Promise<UserRow> {
  if (!UUID_RE.test(id)) throw notFound('User not found.');
  const [u] = await db.select().from(schema.users).where(eq(schema.users.id, id)).limit(1);
  if (!u) throw notFound('User not found.');
  return u;
}

async function adminUser(id: string): Promise<AdminUser> {
  const u = await loadUser(id);
  return toAdminUser(u, (await grantsFor([u.id])).get(u.id) ?? []);
}

/** De-duplicates the ids and confirms each sender account exists. */
async function checkSenderIds(ids: string[]): Promise<string[]> {
  const unique = [...new Set(ids.map(i => i.toLowerCase()))];
  if (!unique.length) return unique;
  const found = await db.select({ id: schema.senderAccounts.id }).from(schema.senderAccounts)
    .where(inArray(schema.senderAccounts.id, unique));
  if (found.length !== unique.length) throw badRequest('One or more of the selected sender accounts no longer exist. Refresh and try again.');
  return unique;
}

const isUniqueViolation = (e: unknown) => {
  const err = e as { code?: string; cause?: { code?: string } };
  return err?.code === '23505' || err?.cause?.code === '23505';
};

const r = new Hono<AppEnv>();

r.get('/', async (c) => {
  const rows = await db.select().from(schema.users).orderBy(asc(schema.users.name));
  const grants = await grantsFor(rows.map(u => u.id));
  return c.json<{ users: AdminUser[] }>({ users: rows.map(u => toAdminUser(u, grants.get(u.id) ?? [])) });
});

r.post('/', async (c) => {
  const actor = me(c);
  const input = await body(c, CreateSchema);
  const [existing] = await db.select({ id: schema.users.id }).from(schema.users).where(eq(schema.users.email, input.email)).limit(1);
  if (existing) throw conflict('A user with this email address already exists.');
  const senderIds = await checkSenderIds(input.senderAccountIds ?? []);

  const temp = tempPassword();
  let created: UserRow;
  try {
    [created] = await db.insert(schema.users).values({
      email: input.email,
      name: input.name,
      role: input.role,
      passwordHash: await hashPassword(temp),
      mustChangePassword: true,
    }).returning();
  } catch (e) {
    if (isUniqueViolation(e)) throw conflict('A user with this email address already exists.');
    throw e;
  }
  if (senderIds.length) {
    await db.insert(schema.userSenderAccess).values(
      senderIds.map(senderAccountId => ({ userId: created.id, senderAccountId, grantedBy: actor.id })),
    );
  }
  await audit(c, 'user.create', {
    targetType: 'user', targetId: created.id,
    meta: { email: created.email, name: created.name, role: created.role, senderAccountIds: senderIds },
  });
  return c.json<{ user: AdminUser; tempPassword: string }>({ user: toAdminUser(created, senderIds), tempPassword: temp }, 201);
});

r.patch('/:id', async (c) => {
  const actor = me(c);
  const target = await loadUser(c.req.param('id'));
  const input = await body(c, UpdateSchema);
  const isSelf = target.id === actor.id;

  if (isSelf && input.disabled === true) throw badRequest('You cannot disable your own account.');
  if (isSelf && input.role === 'sender' && target.role === 'admin') throw badRequest('You cannot remove your own administrator role.');

  const losesAdmin = target.role === 'admin' && !target.disabled
    && (input.role === 'sender' || input.disabled === true);
  if (losesAdmin) {
    const others = await db.select({ id: schema.users.id }).from(schema.users)
      .where(and(eq(schema.users.role, 'admin'), eq(schema.users.disabled, false), ne(schema.users.id, target.id)))
      .limit(1);
    if (!others.length) throw conflict('This is the last active administrator. Make another user an administrator first.');
  }

  const changes: Record<string, { from: unknown; to: unknown }> = {};
  const set: Partial<typeof schema.users.$inferInsert> = {};
  if (input.name !== undefined && input.name !== target.name) { set.name = input.name; changes.name = { from: target.name, to: input.name }; }
  if (input.role !== undefined && input.role !== target.role) { set.role = input.role; changes.role = { from: target.role, to: input.role }; }
  const disabledChanged = input.disabled !== undefined && input.disabled !== target.disabled;
  if (disabledChanged) set.disabled = input.disabled;

  if (Object.keys(set).length) {
    set.updatedAt = new Date();
    await db.update(schema.users).set(set).where(eq(schema.users.id, target.id));
  }
  if (disabledChanged && input.disabled) await revokeUserSessions(target.id);
  // A role change takes effect immediately; signing the user out avoids stale admin screens.
  if (changes.role && !isSelf) await revokeUserSessions(target.id);

  if (Object.keys(changes).length) {
    await audit(c, 'user.update', { targetType: 'user', targetId: target.id, meta: { email: target.email, changes } });
  }
  if (disabledChanged) {
    await audit(c, input.disabled ? 'user.disable' : 'user.enable', { targetType: 'user', targetId: target.id, meta: { email: target.email } });
  }
  return c.json<{ user: AdminUser }>({ user: await adminUser(target.id) });
});

r.post('/:id/reset-password', async (c) => {
  const actor = me(c);
  const target = await loadUser(c.req.param('id'));
  if (target.id === actor.id) throw badRequest('To change your own password, use "Change password" in the account menu.');
  const temp = tempPassword();
  await db.update(schema.users)
    .set({ passwordHash: await hashPassword(temp), mustChangePassword: true, updatedAt: new Date() })
    .where(eq(schema.users.id, target.id));
  await revokeUserSessions(target.id);
  await audit(c, 'user.reset_password', { targetType: 'user', targetId: target.id, meta: { email: target.email } });
  return c.json<{ tempPassword: string }>({ tempPassword: temp });
});

r.put('/:id/senders', async (c) => {
  const actor = me(c);
  const target = await loadUser(c.req.param('id'));
  const { senderAccountIds } = await body(c, SendersSchema);
  const ids = await checkSenderIds(senderAccountIds);
  const before = (await grantsFor([target.id])).get(target.id) ?? [];

  await db.transaction(async (tx) => {
    await tx.delete(schema.userSenderAccess).where(eq(schema.userSenderAccess.userId, target.id));
    if (ids.length) {
      await tx.insert(schema.userSenderAccess).values(ids.map(senderAccountId => ({ userId: target.id, senderAccountId, grantedBy: actor.id })));
    }
  });
  await audit(c, 'user.senders_set', {
    targetType: 'user', targetId: target.id,
    meta: {
      email: target.email,
      senderAccountIds: ids,
      added: ids.filter(i => !before.includes(i)),
      removed: before.filter(i => !ids.includes(i)),
    },
  });
  return c.json<{ user: AdminUser }>({ user: await adminUser(target.id) });
});

export default r;
