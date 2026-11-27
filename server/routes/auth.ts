// /api/auth: sign in, sign out, who am I, change password.
// There is deliberately no sign-up route: accounts are created by an administrator.
import { Hono } from 'hono';
import { and, desc, eq, gt, lt } from 'drizzle-orm';
import { z } from 'zod';
import { db, schema } from '../db/client.js';
import type { AppEnv, UserRow } from '../types.js';
import { badRequest, body, clientIp, iso, tooMany, unauthorized } from '../lib/http.js';
import { createSession, destroySession, me, requireUser, revokeUserSessions } from '../lib/session.js';
import { audit } from '../lib/audit.js';
import { hashPassword, validatePasswordPolicy, verifyDummy, verifyPassword } from '../lib/passwords.js';
import type { MeResponse, PublicUser } from '../../shared/types.js';

export function toPublicUser(u: UserRow): PublicUser {
  return {
    id: u.id,
    email: u.email,
    name: u.name,
    role: u.role,
    mustChangePassword: u.mustChangePassword,
    disabled: u.disabled,
    lastLoginAt: iso(u.lastLoginAt),
    createdAt: u.createdAt.toISOString(),
  };
}

// ------------------------------------------------------------------ rate limiting
const WINDOW_MS = 15 * 60 * 1000;
/** Failed attempts allowed per email (since its last successful sign-in) within the window. */
const MAX_FAILURES_PER_EMAIL = 5;
/**
 * Failed attempts allowed per client IP within the window. Deliberately much higher than the
 * per-email limit: KIIT staff on the campus network share a few public IPs, so a low per-IP
 * cap would let one person's typos lock everyone out. This still stops password spraying.
 */
const MAX_FAILURES_PER_IP = 30;
const RATE_LIMIT_MESSAGE =
  'Too many unsuccessful sign-in attempts. Please wait 15 minutes and try again, or ask the administrator to reset your password.';

async function isRateLimited(email: string, ip: string | null): Promise<boolean> {
  const since = new Date(Date.now() - WINDOW_MS);
  const byEmail = await db.select({ success: schema.loginAttempts.success }).from(schema.loginAttempts)
    .where(and(eq(schema.loginAttempts.email, email), gt(schema.loginAttempts.at, since)))
    .orderBy(desc(schema.loginAttempts.at))
    .limit(50);
  let emailFailures = 0;
  for (const a of byEmail) { if (a.success) break; emailFailures++; }
  if (emailFailures >= MAX_FAILURES_PER_EMAIL) return true;
  if (!ip) return false;
  const byIp = await db.select({ success: schema.loginAttempts.success }).from(schema.loginAttempts)
    .where(and(eq(schema.loginAttempts.ip, ip), eq(schema.loginAttempts.success, false), gt(schema.loginAttempts.at, since)))
    .limit(MAX_FAILURES_PER_IP);
  return byIp.length >= MAX_FAILURES_PER_IP;
}

async function recordAttempt(email: string, ip: string | null, success: boolean): Promise<void> {
  await db.insert(schema.loginAttempts).values({ email, ip, success });
  // Occasional housekeeping: attempts older than a day are no longer useful.
  if (Math.random() < 0.05) {
    await db.delete(schema.loginAttempts).where(lt(schema.loginAttempts.at, new Date(Date.now() - 24 * 60 * 60 * 1000)))
      .catch(() => undefined);
  }
}

// ------------------------------------------------------------------ routes
const LoginSchema = z.object({
  email: z.string().trim().min(1, 'Enter your email address.').max(320).transform(s => s.toLowerCase()),
  password: z.string().min(1, 'Enter your password.').max(1024),
});

const ChangePasswordSchema = z.object({
  currentPassword: z.string().min(1, 'Enter your current password.').max(1024),
  newPassword: z.string().min(1, 'Enter a new password.').max(1024),
});

const GENERIC_LOGIN_ERROR = 'Incorrect email or password.';

const r = new Hono<AppEnv>();

r.post('/login', async (c) => {
  const { email, password } = await body(c, LoginSchema);
  const ip = clientIp(c);

  if (await isRateLimited(email, ip)) {
    await audit(c, 'auth.login_failed', { actorId: null, actorEmail: email, meta: { reason: 'rate_limited' } });
    throw tooMany(RATE_LIMIT_MESSAGE);
  }

  const [user] = await db.select().from(schema.users).where(eq(schema.users.email, email)).limit(1);
  let ok = false;
  if (user) ok = await verifyPassword(user.passwordHash, password);
  else await verifyDummy(password);

  if (!user || !ok || user.disabled) {
    await recordAttempt(email, ip, false);
    await audit(c, 'auth.login_failed', {
      actorId: null, actorEmail: email,
      targetType: user ? 'user' : undefined, targetId: user?.id,
      meta: { reason: !user ? 'unknown_email' : !ok ? 'bad_password' : 'disabled' },
    });
    throw unauthorized(GENERIC_LOGIN_ERROR);
  }

  await recordAttempt(email, ip, true);
  await createSession(c, user.id);
  const now = new Date();
  await db.update(schema.users).set({ lastLoginAt: now }).where(eq(schema.users.id, user.id));
  await audit(c, 'auth.login', { actorId: user.id, actorEmail: user.email, targetType: 'user', targetId: user.id });
  return c.json<MeResponse>({ user: toPublicUser({ ...user, lastLoginAt: now }) });
});

r.post('/logout', async (c) => {
  const user = c.get('user');
  await destroySession(c);
  if (user) await audit(c, 'auth.logout', { targetType: 'user', targetId: user.id });
  return c.json({ ok: true as const });
});

r.get('/me', (c) => {
  const user = c.get('user');
  if (!user) throw unauthorized();
  return c.json<MeResponse>({ user: toPublicUser(user) });
});

r.post('/change-password', requireUser, async (c) => {
  const user = me(c);
  const { currentPassword, newPassword } = await body(c, ChangePasswordSchema);
  const ip = clientIp(c);

  if (await isRateLimited(user.email, ip)) throw tooMany(RATE_LIMIT_MESSAGE);
  if (!(await verifyPassword(user.passwordHash, currentPassword))) {
    await recordAttempt(user.email, ip, false);
    throw badRequest('Your current password is incorrect.');
  }
  if (newPassword === currentPassword) throw badRequest('Choose a new password that is different from your current one.');
  const policyError = validatePasswordPolicy(newPassword, user.email);
  if (policyError) throw badRequest(policyError);

  await db.update(schema.users)
    .set({ passwordHash: await hashPassword(newPassword), mustChangePassword: false, updatedAt: new Date() })
    .where(eq(schema.users.id, user.id));
  await revokeUserSessions(user.id, c.get('sessionTokenHash'));
  await audit(c, 'auth.password_changed', {
    targetType: 'user', targetId: user.id, meta: { wasTemporary: user.mustChangePassword },
  });
  return c.json({ ok: true as const });
});

export default r;
