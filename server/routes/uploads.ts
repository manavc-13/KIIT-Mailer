// /api/uploads: attachments (private) and in-body images. See shared/types.ts.
// Every route needs a signed-in user EXCEPT POST /blob-token when it carries
// Vercel's signed `blob.upload-completed` callback.
import { Hono } from 'hono';
import { z } from 'zod';
import { and, eq } from 'drizzle-orm';
import { handleUpload, type HandleUploadBody } from '@vercel/blob/client';
import { head } from '@vercel/blob';
import type { AppEnv } from '../types.js';
import { db, schema } from '../db/client.js';
import { requireUser, me } from '../lib/session.js';
import { body, badRequest, forbidden, notFound, unauthorized } from '../lib/http.js';
import {
  uploadConfig, assertUploadsAvailable, storageProvider, privateToken, publicToken, maxBytesFor, IMAGE_TYPES, isImageType,
  blobStoreOf, toUploadRef, storeUpload, readUpload, deleteUploadStorage, safeFilename,
} from '../lib/attachments.js';
import { fetchImageSafely, FetchBlockedError } from '../lib/send/ssrf.js';
import type { UploadKind } from '../../shared/types.js';

const r = new Hono<AppEnv>();

const BLOB_TOKEN_PATH = '/api/uploads/blob-token';
r.use('*', async (c, next) => (c.req.path === BLOB_TOKEN_PATH ? next() : requireUser(c, next)));

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function uploadId(raw: string): string {
  if (!UUID_RE.test(raw)) throw notFound('Upload not found.');
  return raw.toLowerCase();
}
const kindSchema = z.enum(['attachment', 'image']);

// ---------------------------------------------------------------- config
r.get('/config', (c) => c.json(uploadConfig()));

// ---------------------------------------------------------------- Vercel Blob client-upload handshake
r.post('/blob-token', async (c) => {
  let payload: HandleUploadBody;
  try { payload = await c.req.json(); } catch { throw badRequest('Request body must be JSON.'); }

  if (payload?.type === 'blob.upload-completed') {
    // Server-to-server callback from Vercel Blob (no session). It is signed with the
    // token of the store that received the upload, so try each configured store.
    for (const token of [privateToken(), publicToken()].filter(Boolean)) {
      try {
        const res = await handleUpload({
          token, request: c.req.raw, body: payload,
          onBeforeGenerateToken: async () => { throw new Error('unexpected'); },
          onUploadCompleted: async ({ blob }) => { console.log('[uploads] blob upload completed', blob.pathname); },
        });
        return c.json(res);
      } catch { /* wrong store or bad signature: try the next */ }
    }
    throw badRequest('Invalid upload callback.');
  }

  if (payload?.type !== 'blob.generate-client-token') throw badRequest('Unknown upload request.');
  const user = c.get('user');
  if (!user) throw unauthorized();
  if (user.mustChangePassword) throw forbidden('Please change your temporary password first.');
  assertUploadsAvailable();
  if (storageProvider() !== 'blob') throw badRequest('This server stores uploads locally. Upload with POST /api/uploads/local instead.');

  let kind: UploadKind = 'attachment';
  try {
    const parsed = JSON.parse(payload.payload?.clientPayload || '{}') as { kind?: string };
    if (parsed.kind === 'image') kind = 'image';
  } catch { /* default attachment */ }

  const usePublic = kind === 'image' && !!publicToken();
  try {
    const res = await handleUpload({
      token: usePublic ? publicToken() : privateToken(),
      request: c.req.raw,
      body: payload,
      onBeforeGenerateToken: async () => ({
        allowedContentTypes: kind === 'image' ? IMAGE_TYPES : undefined,
        maximumSizeInBytes: maxBytesFor(kind),
        addRandomSuffix: true,
        tokenPayload: JSON.stringify({ userId: user.id, kind }),
      }),
      onUploadCompleted: async ({ blob }) => { console.log('[uploads] blob upload completed', blob.pathname); },
    });
    return c.json(res);
  } catch (err) {
    throw badRequest((err as Error).message || 'Could not start the upload.');
  }
});

// ---------------------------------------------------------------- register a finished Blob upload
const registerSchema = z.object({
  kind: kindSchema,
  url: z.string().url().max(2000),
  filename: z.string().min(1).max(300),
  contentType: z.string().min(1).max(200),
  size: z.number().int().nonnegative(),
});

r.post('/register', async (c) => {
  const user = me(c);
  assertUploadsAvailable();
  const b = await body(c, registerSchema);
  const store = blobStoreOf(b.url);
  if (!store) throw badRequest('That file URL is not from this app\'s Blob storage.');
  if (b.kind === 'attachment' && store !== 'private') throw badRequest('Attachments must be stored in the private Blob store.');

  let size = b.size; let contentType = b.contentType;
  try {
    const meta = await head(b.url, { token: store === 'public' ? publicToken() : privateToken() });
    size = meta.size; contentType = meta.contentType || contentType;
  } catch {
    throw badRequest('The uploaded file could not be found in storage. Please upload it again.');
  }
  if (size > maxBytesFor(b.kind)) throw badRequest(`The file is larger than the ${(maxBytesFor(b.kind) / 1024 / 1024).toFixed(1)} MB limit.`);
  if (b.kind === 'image' && !isImageType(contentType)) throw badRequest('Images must be PNG, JPEG, GIF or WebP.');

  const [row] = await db.insert(schema.uploads).values({
    ownerId: user.id, kind: b.kind, provider: 'blob', url: b.url, localId: null,
    filename: safeFilename(b.filename), contentType, size,
  }).returning();
  return c.json({ upload: toUploadRef(row) });
});

// ---------------------------------------------------------------- local dev provider
r.post('/local', async (c) => {
  const user = me(c);
  if (storageProvider() !== 'local') throw badRequest('Local uploads are only available on the local development server.');
  const kind = kindSchema.safeParse(c.req.query('kind') || 'attachment');
  if (!kind.success) throw badRequest('kind must be "attachment" or "image".');
  const filename = c.req.query('filename') || (kind.data === 'image' ? 'image' : 'attachment');
  const contentType = c.req.query('contentType') || c.req.header('content-type') || 'application/octet-stream';
  const declared = Number(c.req.header('content-length') || 0);
  if (declared > maxBytesFor(kind.data)) throw badRequest(`The file is larger than the ${(maxBytesFor(kind.data) / 1024 / 1024).toFixed(1)} MB limit.`);
  const buf = Buffer.from(await c.req.arrayBuffer());
  const row = await storeUpload(user.id, kind.data, buf, filename, contentType);
  return c.json({ upload: toUploadRef(row) });
});

// ---------------------------------------------------------------- rehost a pasted image URL
r.post('/image-from-url', async (c) => {
  const user = me(c);
  assertUploadsAvailable();
  const b = await body(c, z.object({ url: z.string().min(1).max(4000) }));
  let img;
  try { img = await fetchImageSafely(b.url); } catch (err) {
    if (err instanceof FetchBlockedError) throw badRequest(err.message);
    throw badRequest('The image could not be downloaded.');
  }
  const row = await storeUpload(user.id, 'image', img.buffer, img.filename, img.contentType);
  return c.json({ upload: toUploadRef(row) });
});

// ---------------------------------------------------------------- read / delete
const INLINE_TYPES = [...IMAGE_TYPES, 'application/pdf'];

r.get('/:id/content', async (c) => {
  const user = me(c);
  const id = uploadId(c.req.param('id'));
  const [row] = await db.select().from(schema.uploads).where(eq(schema.uploads.id, id)).limit(1);
  // Images may appear in shared templates, so any signed-in user can view them; files only their owner.
  if (!row || (row.ownerId !== user.id && row.kind !== 'image')) throw notFound('Upload not found.');
  let buf: Buffer;
  try { buf = await readUpload(row); } catch (err) { throw notFound((err as Error).message); }
  const ct = row.contentType.split(';')[0].trim().toLowerCase();
  const inline = INLINE_TYPES.includes(ct);
  const fname = row.filename.replace(/[^\x20-\x7e]/g, '_').replace(/"/g, '');
  return c.body(new Uint8Array(buf), 200, {
    'Content-Type': row.contentType || 'application/octet-stream',
    'Content-Length': String(buf.length),
    'Content-Disposition': `${inline ? 'inline' : 'attachment'}; filename="${fname}"; filename*=UTF-8''${encodeURIComponent(row.filename)}`,
    'X-Content-Type-Options': 'nosniff',
    'Content-Security-Policy': "default-src 'none'; img-src 'self' data:; style-src 'unsafe-inline'; sandbox",
  });
});

r.delete('/:id', async (c) => {
  const user = me(c);
  const id = uploadId(c.req.param('id'));
  const [row] = await db.select().from(schema.uploads).where(and(eq(schema.uploads.id, id), eq(schema.uploads.ownerId, user.id))).limit(1);
  if (!row) throw notFound('Upload not found.');
  // Files used by a send job are kept: history and "retry failed" still need them.
  const [used] = await db.select({ j: schema.jobAttachments.jobId }).from(schema.jobAttachments).where(eq(schema.jobAttachments.uploadId, id)).limit(1);
  if (used) return c.json({ ok: true as const, kept: true });
  await deleteUploadStorage(row);
  await db.delete(schema.uploads).where(eq(schema.uploads.id, id));
  return c.json({ ok: true as const });
});

export default r;
