// Upload storage and resolution (ported from legacy/lib/attachments.js and
// legacy/lib/local-upload.js, now owner-checked and DB-backed).
//
// Providers:
//   'blob'  - on Vercel with BLOB_READ_WRITE_TOKEN (a PRIVATE store). The browser
//             uploads straight to Blob storage; the server reads it back with get().
//   'local' - local dev server: files in os.tmpdir()/kiit-mailer-uploads.
// In-body images go to the PUBLIC store (BLOB_PUBLIC_READ_WRITE_TOKEN) when it exists;
// otherwise they are stored like attachments and turned into CID inline images at send.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { randomUUID } from 'node:crypto';
import { and, eq, inArray } from 'drizzle-orm';
import { get, put, del } from '@vercel/blob';
import { db, schema } from '../db/client.js';
import { badRequest } from './http.js';
import {
  GMAIL_MAX_MESSAGE_BYTES, MIME_OVERHEAD_BYTES, SOFT_WARN_BYTES, MAX_IMAGE_BYTES, maxRawAttachmentBytes,
  type MailAttachment,
} from './mailer.js';
import type { HostedImage } from './send/images.js';
import type { UploadConfig, UploadKind, UploadRef } from '../../shared/types.js';

export type UploadRow = typeof schema.uploads.$inferSelect;

export const LOCAL_UPLOAD_DIR = path.join(os.tmpdir(), 'kiit-mailer-uploads');

export const isVercel = () => !!process.env.VERCEL;
export const privateToken = () => process.env.BLOB_READ_WRITE_TOKEN || '';
export const publicToken = () => process.env.BLOB_PUBLIC_READ_WRITE_TOKEN || '';

export function storageProvider(): 'blob' | 'local' {
  return isVercel() ? 'blob' : 'local';
}

/** Null when uploads work; otherwise a plain-English reason (and uploads answer 400). */
export function degradedReason(): string | null {
  if (isVercel() && !privateToken()) {
    return 'No Vercel Blob store is connected to this deployment, so files cannot be uploaded. An administrator needs to create a private Blob store and set BLOB_READ_WRITE_TOKEN.';
  }
  return null;
}

export function uploadConfig(): UploadConfig {
  return {
    provider: storageProvider(),
    publicImages: !!publicToken(),
    maxRawAttachmentBytes: maxRawAttachmentBytes(),
    maxEncodedMessageBytes: GMAIL_MAX_MESSAGE_BYTES,
    softWarnBytes: SOFT_WARN_BYTES,
    mimeOverheadBytes: MIME_OVERHEAD_BYTES,
    degradedReason: degradedReason(),
  };
}

export function assertUploadsAvailable(): void {
  const reason = degradedReason();
  if (reason) throw badRequest(reason);
}

export const maxBytesFor = (kind: UploadKind) => (kind === 'image' ? MAX_IMAGE_BYTES : maxRawAttachmentBytes());

export const IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/webp'];
export const isImageType = (ct: string) => IMAGE_TYPES.includes(ct.split(';')[0].trim().toLowerCase().replace('image/jpg', 'image/jpeg'));

/** Store id embedded in a Blob read-write token ("vercel_blob_rw_<storeId>_<secret>"). */
export function blobStoreId(token: string): string | null {
  const m = /^vercel_blob_rw_([A-Za-z0-9]+)_/.exec(token || '');
  return m ? m[1].toLowerCase() : null;
}

/** Which configured store a blob URL belongs to (null = not ours / not a blob URL). */
export function blobStoreOf(url: string): 'private' | 'public' | null {
  let host: string;
  try {
    const u = new URL(url);
    if (u.protocol !== 'https:') return null;
    host = u.hostname.toLowerCase();
  } catch { return null; }
  if (!host.endsWith('.blob.vercel-storage.com')) return null;
  const matches = (tok: string) => {
    if (!tok) return false;
    const id = blobStoreId(tok);
    return id ? host.startsWith(id + '.') : true; // unknown token format: accept by suffix only
  };
  if (matches(publicToken()) && (host.includes('.public.') || !matches(privateToken()))) return 'public';
  if (matches(privateToken())) return 'private';
  return null;
}

export function publicUrlFor(row: Pick<UploadRow, 'id' | 'kind' | 'url'>): string | null {
  if (row.kind !== 'image') return null;
  if (row.url && blobStoreOf(row.url) === 'public') return row.url;
  return `/api/uploads/${row.id}/content`;
}

export function toUploadRef(row: UploadRow): UploadRef {
  return { id: row.id, kind: row.kind, filename: row.filename, contentType: row.contentType, size: row.size, publicUrl: publicUrlFor(row) };
}

export const safeFilename = (name: string, fallback = 'attachment') =>
  (String(name || '').replace(/[\\/\0\r\n"]+/g, '_').trim().slice(0, 200)) || fallback;

// ---------------------------------------------------------------- local provider
function ensureLocalDir(): string {
  if (!fs.existsSync(LOCAL_UPLOAD_DIR)) fs.mkdirSync(LOCAL_UPLOAD_DIR, { recursive: true });
  return LOCAL_UPLOAD_DIR;
}
const localPath = (id: string) => path.join(ensureLocalDir(), path.basename(id));
const localMetaPath = (id: string) => localPath(id) + '.meta.json';

async function writeLocal(buf: Buffer, meta: { filename: string; contentType: string }): Promise<string> {
  const id = randomUUID();
  await fs.promises.writeFile(localPath(id), buf);
  await fs.promises.writeFile(localMetaPath(id), JSON.stringify({ ...meta, size: buf.length }));
  return id;
}

// ---------------------------------------------------------------- create
/** Store bytes server-side (local file, or Blob put) and insert the uploads row. */
export async function storeUpload(ownerId: string, kind: UploadKind, buf: Buffer, filename: string, contentType: string): Promise<UploadRow> {
  assertUploadsAvailable();
  if (!buf.length) throw badRequest('The file is empty.');
  if (buf.length > maxBytesFor(kind)) throw badRequest(`The file is larger than the ${(maxBytesFor(kind) / 1024 / 1024).toFixed(1)} MB limit.`);
  if (kind === 'image' && !isImageType(contentType)) throw badRequest('Images must be PNG, JPEG, GIF or WebP.');
  const name = safeFilename(filename, kind === 'image' ? 'image' : 'attachment');

  let provider: string; let url: string | null = null; let localId: string | null = null;
  if (kind === 'image' && publicToken()) {
    const r = await put(`images/${name}`, buf, { access: 'public', token: publicToken(), contentType, addRandomSuffix: true });
    provider = 'blob'; url = r.url;
  } else if (storageProvider() === 'blob') {
    const r = await put(`${kind === 'image' ? 'images' : 'attachments'}/${name}`, buf, { access: 'private', token: privateToken(), contentType, addRandomSuffix: true });
    provider = 'blob'; url = r.url;
  } else {
    provider = 'local'; localId = await writeLocal(buf, { filename: name, contentType });
  }
  const [row] = await db.insert(schema.uploads).values({
    ownerId, kind, provider, url, localId, filename: name, contentType, size: buf.length,
  }).returning();
  return row;
}

// ---------------------------------------------------------------- read
/** Read an upload's bytes (private/public blob or local file). */
export async function readUpload(row: UploadRow): Promise<Buffer> {
  const gone = () => new Error(`"${row.filename}" is no longer available (it may have been deleted or expired). Please upload it again.`);
  if (row.provider === 'local') {
    if (!row.localId) throw gone();
    try { return await fs.promises.readFile(localPath(row.localId)); } catch { throw gone(); }
  }
  if (!row.url) throw gone();
  const store = blobStoreOf(row.url);
  if (store === 'public') {
    const res = await fetch(row.url);
    if (!res.ok) throw gone();
    return Buffer.from(await res.arrayBuffer());
  }
  // Private store: a plain fetch gets 400 ("Cannot use public access on a private store"); use the SDK.
  const result = await get(row.url, { access: 'private', token: privateToken() || undefined });
  if (!result || result.statusCode !== 200 || !result.stream) throw gone();
  const chunks: Uint8Array[] = [];
  for await (const c of result.stream as unknown as AsyncIterable<Uint8Array>) chunks.push(c);
  return Buffer.concat(chunks);
}

/** Remove the stored bytes (best effort). The DB row is handled by the caller. */
export async function deleteUploadStorage(row: UploadRow): Promise<void> {
  try {
    if (row.provider === 'local' && row.localId) {
      await fs.promises.rm(localPath(row.localId), { force: true });
      await fs.promises.rm(localMetaPath(row.localId), { force: true });
    } else if (row.url) {
      const store = blobStoreOf(row.url);
      const token = store === 'public' ? publicToken() : privateToken();
      if (token) await del(row.url, { token });
    }
  } catch (err) {
    console.error('[uploads] failed to delete stored file', row.id, err);
  }
}

// ---------------------------------------------------------------- send-time resolution
/**
 * Resolve attachment upload ids into buffers ONCE per run; reused for every message.
 * Every upload must exist, belong to `ownerId` and be kind 'attachment'.
 */
export async function resolveAttachments(uploadIds: string[], ownerId: string): Promise<MailAttachment[]> {
  const ids = [...new Set(uploadIds)];
  if (!ids.length) return [];
  const rows = await db.select().from(schema.uploads)
    .where(and(inArray(schema.uploads.id, ids), eq(schema.uploads.ownerId, ownerId), eq(schema.uploads.kind, 'attachment')));
  if (rows.length !== ids.length) throw new Error('One or more attachments were not found or are not yours. Please attach them again.');
  const byId = new Map(rows.map(r => [r.id, r]));
  return Promise.all(ids.map(async id => {
    const r = byId.get(id)!;
    return { filename: r.filename, content: await readUpload(r), contentType: r.contentType };
  }));
}

/**
 * Load app-hosted images (kind 'image', any owner: images live in shared templates)
 * for CID embedding. Missing ones are skipped (the URL stays as-is).
 */
export async function loadHostedImages(uploadIds: string[]): Promise<HostedImage[]> {
  const ids = [...new Set(uploadIds)];
  if (!ids.length) return [];
  const rows = await db.select().from(schema.uploads).where(and(inArray(schema.uploads.id, ids), eq(schema.uploads.kind, 'image')));
  const out: HostedImage[] = [];
  for (const r of rows) {
    try { out.push({ id: r.id, filename: r.filename, contentType: r.contentType, content: await readUpload(r) }); }
    catch (err) { console.error('[uploads] hosted image unavailable', r.id, (err as Error).message); }
  }
  return out;
}

/** Validates attachment ids for a new job: owned by the user, kind attachment, total within budget. */
export async function assertOwnedAttachments(uploadIds: string[], ownerId: string): Promise<UploadRow[]> {
  const ids = [...new Set(uploadIds)];
  if (!ids.length) return [];
  const rows = await db.select().from(schema.uploads)
    .where(and(inArray(schema.uploads.id, ids), eq(schema.uploads.ownerId, ownerId), eq(schema.uploads.kind, 'attachment')));
  if (rows.length !== ids.length) throw badRequest('One or more attachments were not found or are not yours. Please attach them again.');
  const total = rows.reduce((s, r) => s + r.size, 0);
  if (total > maxRawAttachmentBytes()) {
    throw badRequest(`Attachments total ${(total / 1024 / 1024).toFixed(1)} MB, over the ${(maxRawAttachmentBytes() / 1024 / 1024).toFixed(1)} MB limit for one email.`);
  }
  return rows;
}
