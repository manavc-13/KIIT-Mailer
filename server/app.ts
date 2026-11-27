// The whole API as one Hono app, mounted at /api. Runs as a single Vercel
// function (api/[[...route]].ts) and locally via server/dev.ts.
//
// Guards: loadSession runs for every request and sets c.get('user').
// /api/admin/* requires an admin here. Every other route file applies
// requireUser itself (so public paths like /api/u/* and the Blob upload
// callback can opt out).
import { Hono } from 'hono';
import { HTTPException } from 'hono/http-exception';
import type { AppEnv } from './types.js';
import { ApiError } from './lib/http.js';
import { loadSession, requireAdmin, csrf } from './lib/session.js';
import auth from './routes/auth.js';
import senders from './routes/senders.js';
import templates from './routes/templates.js';
import uploads from './routes/uploads.js';
import render from './routes/render.js';
import jobs from './routes/jobs.js';
import unsubscribe from './routes/unsubscribe.js';
import adminUsers from './routes/admin-users.js';
import adminSenders from './routes/admin-senders.js';
import adminAudit from './routes/admin-audit.js';
import adminSuppressions from './routes/admin-suppressions.js';

const CSRF_EXEMPT = (p: string) => p.startsWith('/api/u/') || p === '/api/uploads/blob-token';

const app = new Hono<AppEnv>().basePath('/api');

app.use('*', async (c, next) => {
  await next();
  c.header('Cache-Control', 'no-store');
});
app.use('*', loadSession);
app.use('*', csrf(CSRF_EXEMPT));

app.get('/health', (c) => c.json({ ok: true }));

app.route('/auth', auth);
app.route('/u', unsubscribe);
app.route('/senders', senders);
app.route('/templates', templates);
app.route('/uploads', uploads);
app.route('/render', render);
app.route('/jobs', jobs);

app.use('/admin/*', requireAdmin);
app.route('/admin/users', adminUsers);
app.route('/admin/senders', adminSenders);
app.route('/admin/audit', adminAudit);
app.route('/admin/suppressions', adminSuppressions);

app.notFound((c) => c.json({ error: 'Not found.', code: 'not_found' }, 404));

app.onError((err, c) => {
  if (err instanceof ApiError) {
    return c.json({ error: err.message, code: err.code, details: err.details }, err.status);
  }
  if (err instanceof HTTPException) {
    return c.json({ error: err.message || 'Request failed.' }, err.status);
  }
  console.error('[api] unhandled error', c.req.method, c.req.path, err);
  return c.json({ error: 'Something went wrong on the server. Please try again.', code: 'internal' }, 500);
});

export default app;

/**
 * Entry used by the servers (server/dev.ts, server/vercel.ts). It is app.fetch plus one alias:
 * unsubscribe links printed in emails are short `/u/<token>` URLs, which are served by the
 * `/api/u` routes.
 */
export function handleRequest(req: Request, ...rest: unknown[]): Response | Promise<Response> {
  const url = new URL(req.url);
  if (url.pathname.startsWith('/u/')) {
    url.pathname = '/api' + url.pathname;
    const init: RequestInit & { duplex?: 'half' } = { method: req.method, headers: req.headers };
    if (req.body) { init.body = req.body; init.duplex = 'half'; }
    req = new Request(url, init);
  }
  return app.fetch(req, ...(rest as []));
}
