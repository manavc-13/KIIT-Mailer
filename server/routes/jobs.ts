// /api/jobs: create, pump (run), pause/cancel/retry, history, CSV export, test send.
import { Hono } from 'hono';
import { z } from 'zod';
import type { AppEnv } from '../types.js';
import { requireUser, me } from '../lib/session.js';
import { body, query } from '../lib/http.js';
import { toCSV, isValidEmail } from '../../shared/text.js';
import {
  createJob, runJob, pause, cancel, retryFailed, listJobs, activeJobs, jobDetail, jobMessagesForExport, testSend,
  MAX_MESSAGES_PER_JOB,
} from '../lib/jobs.js';
import type { CreateJobRequest, TestSendRequest } from '../../shared/types.js';

const r = new Hono<AppEnv>();
r.use('*', requireUser);

// Request-size limits. Vercel rejects function request bodies over 4.5 MB before our code
// runs (with an unfriendly error), so a create-job request is kept well under that:
// body HTML 1 MB + plain text 200 KB + recipient list 2.5 MB (JSON) + small fields.
const KB = 1024;
const MB = 1024 * KB;
export const MAX_BODY_HTML_BYTES = 1 * MB;
export const MAX_BODY_TEXT_BYTES = 200 * KB;
export const MAX_RECIPIENT_ROWS = MAX_MESSAGES_PER_JOB; // 5,000, same as the per-send email cap
export const MAX_RECIPIENT_LIST_BYTES = 2.5 * MB;
export const MAX_COLUMNS = 100;
export const MAX_COLUMN_NAME_CHARS = 100;
export const MAX_CELL_CHARS = 5000;
const utf8Bytes = (s: string) => Buffer.byteLength(s, 'utf8');

const themeSchema = z.enum(['heritage', 'meridian', 'letterhead', 'slate', 'plain']);
const rowSchema = z.record(
  z.string(),
  z.coerce.string().max(MAX_CELL_CHARS, `A value in the recipient list is too long (over ${MAX_CELL_CHARS.toLocaleString('en-US')} characters).`),
)
  .refine(r => Object.keys(r).length <= MAX_COLUMNS, `The recipient list has too many columns (over ${MAX_COLUMNS}). Remove the ones the email does not use.`)
  .refine(r => Object.keys(r).every(k => k.length <= MAX_COLUMN_NAME_CHARS), `A column name in the recipient list is too long (over ${MAX_COLUMN_NAME_CHARS} characters).`);

export const composeSchema = z.object({
  subject: z.string().max(998, 'The subject is too long.'),
  preheader: z.string().max(500, 'The preheader is too long (over 500 characters).').default(''),
  editorMode: z.enum(['rich', 'html']),
  bodyHtml: z.string().refine(s => utf8Bytes(s) <= MAX_BODY_HTML_BYTES,
    'The email body is too large (over 1 MB). Remove pasted images or long content, or send large files as attachments.'),
  bodyText: z.string().refine(s => utf8Bytes(s) <= MAX_BODY_TEXT_BYTES,
    'The plain-text version of the email is too large (over 200 KB).').nullish(),
  theme: themeSchema,
  includeSignature: z.boolean().default(true),
});

const attachmentIdsSchema = z.array(z.string().uuid()).max(20, 'Attach at most 20 files.').default([]);

const createSchema = composeSchema.extend({
  senderAccountId: z.string().uuid('Choose a sender account.'),
  templateId: z.string().nullish(),
  mode: z.enum(['personal', 'bcc_broadcast']),
  recipients: z.array(z.object({
    to: z.string().max(2000, 'A recipient\'s To field is too long (over 2,000 characters).'),
    cc: z.string().max(5000, 'A recipient\'s CC field is too long (over 5,000 characters).').optional(),
    bcc: z.string().max(10_000, 'A recipient\'s BCC field is too long (over 10,000 characters).').optional(),
    data: rowSchema.optional(),
  }))
    .min(1, 'Add at least one recipient.')
    .max(MAX_RECIPIENT_ROWS, `Too many recipients in one send (the limit is ${MAX_RECIPIENT_ROWS.toLocaleString('en-US')} rows). Split the list into smaller sends.`)
    .refine(list => utf8Bytes(JSON.stringify(list)) <= MAX_RECIPIENT_LIST_BYTES,
      'The recipient list is too large to send in one go (over 2.5 MB of addresses and column data). Remove columns the email does not use, or split the list into smaller sends.'),
  globalCc: z.string().max(5000, 'The CC-everyone field is too long (over 5,000 characters).').optional(),
  globalBcc: z.string().max(20_000, 'The BCC-everyone field is too long (over 20,000 characters).').optional(),
  attachmentIds: attachmentIdsSchema,
  unsubscribe: z.union([z.boolean(), z.literal('auto')]).default('auto'),
  dedupe: z.boolean().optional(),
});

const testSchema = composeSchema.extend({
  senderAccountId: z.string().uuid('Choose a sender account.'),
  to: z.array(z.string().trim().max(320, 'A test address is too long.')).min(1, 'Enter at least one test address.').max(5, 'Send a test to at most 5 addresses.')
    .refine(list => list.every(isValidEmail), 'One of the test addresses is not a valid email address.'),
  row: rowSchema.optional(),
  attachmentIds: attachmentIdsSchema,
});

export { createSchema as createJobSchema, testSchema as testSendSchema };

r.post('/', async (c) => {
  const b = await body(c, createSchema);
  return c.json(await createJob(c, me(c), b as CreateJobRequest));
});

r.post('/test', async (c) => {
  const b = await body(c, testSchema);
  return c.json(await testSend(c, me(c), b as TestSendRequest));
});

r.get('/', async (c) => {
  const q = query(c, z.object({ scope: z.enum(['mine', 'all']).default('mine'), cursor: z.string().optional() }));
  return c.json(await listJobs(me(c), q.scope, q.cursor));
});

r.get('/active', async (c) => c.json({ jobs: await activeJobs(me(c)) }));

r.get('/:id', async (c) => c.json(await jobDetail(me(c), c.req.param('id'))));

r.get('/:id/export.csv', async (c) => {
  const { job, messages } = await jobMessagesForExport(me(c), c.req.param('id'));
  const cols = ['seq', 'to', 'cc', 'bcc', 'recipients', 'status', 'attempts', 'error', 'sent_at', 'message_id'];
  const csv = toCSV(cols, messages.map(m => ({
    seq: m.seq, to: m.to, cc: m.cc ?? '', bcc: m.bcc ?? '', recipients: m.recipientCount, status: m.status,
    attempts: m.attempts, error: m.error ?? '', sent_at: m.sentAt ? m.sentAt.toISOString() : '', message_id: m.smtpMessageId ?? '',
  })));
  const date = job.createdAt.toISOString().slice(0, 10);
  return c.body('﻿' + csv, 200, {
    'Content-Type': 'text/csv; charset=utf-8',
    'Content-Disposition': `attachment; filename="kiit-mailer-send-${date}-${job.id.slice(0, 8)}.csv"`,
  });
});

r.post('/:id/run', async (c) => c.json(await runJob(c, me(c), c.req.param('id'))));
r.post('/:id/pause', async (c) => c.json({ job: await pause(c, me(c), c.req.param('id')) }));
r.post('/:id/cancel', async (c) => c.json({ job: await cancel(c, me(c), c.req.param('id')) }));
r.post('/:id/retry-failed', async (c) => c.json({ job: await retryFailed(c, me(c), c.req.param('id')) }));

export default r;
