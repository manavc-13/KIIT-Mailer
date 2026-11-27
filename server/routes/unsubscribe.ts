// Public one-click unsubscribe (RFC 8058). Mounted at /api/u; vercel.json rewrites /u/* here.
// No session, no CSRF header (exempt in app.ts): the signed token is the authority.
//   GET  /:token  confirmation page with a "Confirm unsubscribe" button
//   POST /:token  suppress the address (form button, or mail client one-click
//                 with body "List-Unsubscribe=One-Click")
import { Hono } from 'hono';
import { eq } from 'drizzle-orm';
import type { AppEnv } from '../types.js';
import { db, schema } from '../db/client.js';
import { audit } from '../lib/audit.js';
import { parseUnsubscribeToken, confirmPage, donePage, invalidPage, previewPage } from '../lib/unsubscribe.js';

const r = new Hono<AppEnv>();

const HTML_HEADERS = { 'Content-Type': 'text/html; charset=utf-8', 'X-Robots-Tag': 'noindex', 'Referrer-Policy': 'no-referrer' };

r.get('/preview', (c) => c.body(previewPage(), 200, HTML_HEADERS));

r.get('/:token', (c) => {
  const p = parseUnsubscribeToken(c.req.param('token'));
  if (!p) return c.body(invalidPage(), 400, HTML_HEADERS);
  return c.body(confirmPage(p.e), 200, HTML_HEADERS);
});

r.post('/:token', async (c) => {
  const p = parseUnsubscribeToken(c.req.param('token'));
  if (!p) return c.body(invalidPage(), 400, HTML_HEADERS);
  const email = p.e.toLowerCase();

  let senderAccountId: string | null = null;
  if (/^[0-9a-f-]{36}$/i.test(p.s)) {
    const [s] = await db.select({ id: schema.senderAccounts.id }).from(schema.senderAccounts).where(eq(schema.senderAccounts.id, p.s)).limit(1);
    senderAccountId = s?.id ?? null;
  }
  const [existing] = await db.select({ reason: schema.suppressions.reason }).from(schema.suppressions).where(eq(schema.suppressions.email, email)).limit(1);
  if (!existing) {
    await db.insert(schema.suppressions)
      .values({ email, reason: 'unsubscribed', senderAccountId, note: p.j ? `One-click unsubscribe (job ${p.j})` : 'One-click unsubscribe' })
      .onConflictDoNothing();
    await audit(c, 'suppression.add', {
      actorId: null, actorEmail: null, targetType: 'suppression', targetId: email,
      meta: { reason: 'unsubscribed', senderAccountId, jobId: p.j ?? null, oneClick: /List-Unsubscribe=One-Click/i.test(await safeText(c.req.raw)) },
    });
  }
  return c.body(donePage(email), 200, HTML_HEADERS);
});

async function safeText(req: Request): Promise<string> {
  try { return (await req.text()).slice(0, 200); } catch { return ''; }
}

export default r;
