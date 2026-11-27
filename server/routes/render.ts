// POST /api/render/preview: exactly the send pipeline (server/lib/render), for one row.
import { Hono } from 'hono';
import { z } from 'zod';
import type { AppEnv } from '../types.js';
import { requireUser, me } from '../lib/session.js';
import { body } from '../lib/http.js';
import { assertSenderAccess } from '../lib/access.js';
import { renderEmail } from '../lib/render/index.js';
import { previewUnsubscribeUrl } from '../lib/unsubscribe.js';
import type { ComposeContent, RenderResult } from '../../shared/types.js';
import { composeSchema } from './jobs.js';

const r = new Hono<AppEnv>();
r.use('*', requireUser);

const previewSchema = composeSchema.extend({
  senderAccountId: z.string().uuid('Choose a sender account.'),
  row: z.record(z.string(), z.coerce.string()).optional(),
  unsubscribe: z.boolean().optional(),
});

r.post('/preview', async (c) => {
  const b = await body(c, previewSchema);
  const sender = await assertSenderAccess(me(c), b.senderAccountId);
  const out = renderEmail({
    content: b as ComposeContent,
    sender,
    row: b.row || {},
    unsubscribeUrl: b.unsubscribe ? previewUnsubscribeUrl() : null,
  });
  const result: RenderResult = {
    subject: out.subject, html: out.html, text: out.text, sizeBytes: out.sizeBytes, lint: out.lint, webFontsUsed: out.webFontsUsed,
  };
  return c.json(result);
});

export default r;
