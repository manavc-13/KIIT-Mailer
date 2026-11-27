// Cookie sessions. The browser holds a random 256-bit token in an httpOnly,
// Secure, SameSite=Lax cookie; the DB stores only sha256(token). Sessions slide
// (12 h idle timeout) up to an absolute 7-day lifetime.
import type { Context, MiddlewareHandler } from 'hono';
import { getCookie, setCookie, deleteCookie } from 'hono/cookie';
import { and, eq, gt } from 'drizzle-orm';
import { db, schema } from '../db/client.js';
import { randomToken, sha256Hex } from './crypto.js';
import { clientIp, forbidden, unauthorized } from './http.js';
import type { AppEnv, UserRow } from '../types.js';

export const SESSION_COOKIE = 'km_session';
const IDLE_MS = 12 * 60 * 60 * 1000;
const ABSOLUTE_MS = 7 * 24 * 60 * 60 * 1000;
const TOUCH_EVERY_MS = 5 * 60 * 1000;

const secure = () => process.env.NODE_ENV === 'production' || !!process.env.VERCEL;

export async function createSession(c: Context<AppEnv>, userId: string): Promise<void> {
  const token = randomToken(32);
  await db.insert(schema.sessions).values({
    tokenHash: sha256Hex(token),
    userId,
    expiresAt: new Date(Date.now() + IDLE_MS),
    ip: clientIp(c),
    userAgent: c.req.header('user-agent')?.slice(0, 300) ?? null,
  });
  setCookie(c, SESSION_COOKIE, token, { httpOnly: true, secure: secure(), sameSite: 'Lax', path: '/', maxAge: ABSOLUTE_MS / 1000 });
}

export async function destroySession(c: Context<AppEnv>): Promise<void> {
  const hash = c.get('sessionTokenHash');
  if (hash) await db.delete(schema.sessions).where(eq(schema.sessions.tokenHash, hash));
  deleteCookie(c, SESSION_COOKIE, { path: '/' });
}

/** Revoke every session of a user (password change/reset, disable). Optionally keep the current one. */
export async function revokeUserSessions(userId: string, exceptTokenHash?: string | null): Promise<void> {
  const rows = await db.select({ h: schema.sessions.tokenHash }).from(schema.sessions).where(eq(schema.sessions.userId, userId));
  for (const r of rows) {
    if (r.h !== exceptTokenHash) await db.delete(schema.sessions).where(eq(schema.sessions.tokenHash, r.h));
  }
}

/** Loads c.get('user') from the session cookie (null if absent/expired/disabled). Never throws 401 itself. */
export const loadSession: MiddlewareHandler<AppEnv> = async (c, next) => {
  c.set('user', null);
  c.set('sessionTokenHash', null);
  const token = getCookie(c, SESSION_COOKIE);
  if (token) {
    const hash = sha256Hex(token);
    const now = new Date();
    const rows = await db
      .select({ s: schema.sessions, u: schema.users })
      .from(schema.sessions)
      .innerJoin(schema.users, eq(schema.users.id, schema.sessions.userId))
      .where(and(eq(schema.sessions.tokenHash, hash), gt(schema.sessions.expiresAt, now)))
      .limit(1);
    const row = rows[0];
    if (row && !row.u.disabled && now.getTime() - row.s.createdAt.getTime() < ABSOLUTE_MS) {
      c.set('user', row.u);
      c.set('sessionTokenHash', hash);
      if (now.getTime() - row.s.lastSeenAt.getTime() > TOUCH_EVERY_MS) {
        await db.update(schema.sessions)
          .set({ lastSeenAt: now, expiresAt: new Date(now.getTime() + IDLE_MS) })
          .where(eq(schema.sessions.tokenHash, hash));
      }
    }
  }
  await next();
};

/** 401 unless signed in. Users who must change their password may only reach /api/auth/*. */
export const requireUser: MiddlewareHandler<AppEnv> = async (c, next) => {
  const user = c.get('user');
  if (!user) throw unauthorized();
  if (user.mustChangePassword && !c.req.path.startsWith('/api/auth/')) {
    throw forbidden('Please change your temporary password first.');
  }
  await next();
};

export const requireAdmin: MiddlewareHandler<AppEnv> = async (c, next) => {
  const user = c.get('user');
  if (!user) throw unauthorized();
  if (user.role !== 'admin') throw forbidden('Administrator access required.');
  if (user.mustChangePassword) throw forbidden('Please change your temporary password first.');
  await next();
};

/** Current user, asserting presence (use after requireUser). */
export function me(c: Context<AppEnv>): UserRow {
  const u = c.get('user');
  if (!u) throw unauthorized();
  return u;
}

/**
 * CSRF defence in depth on top of SameSite=Lax: every state-changing request must
 * carry `X-Requested-With: kiit-mailer` (a custom header a cross-site form cannot set).
 * Paths in `exempt` (public one-click unsubscribe, Vercel Blob server callback) skip it.
 */
export function csrf(exempt: (path: string) => boolean): MiddlewareHandler<AppEnv> {
  return async (c, next) => {
    const m = c.req.method;
    if (m !== 'GET' && m !== 'HEAD' && m !== 'OPTIONS' && !exempt(c.req.path)) {
      if (c.req.header('x-requested-with') !== 'kiit-mailer') throw forbidden('Missing request header.');
    }
    await next();
  };
}
