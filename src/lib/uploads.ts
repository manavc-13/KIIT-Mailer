// Client upload helper, used by the editor (in-body images) and compose (attachments).
//
// Flow (see shared/types.ts "uploads"):
//   provider 'blob'  -> @vercel/blob/client upload() with handleUploadUrl '/api/uploads/blob-token'
//                       (clientPayload: JSON.stringify({ kind }), header X-Requested-With), then
//                       POST /api/uploads/register.
//   provider 'local' -> POST /api/uploads/local?kind=&filename=&contentType= with the raw file body
//                       (XHR so upload progress is reported).
//   ('inline' is not supported in v2: a Blob store is required on Vercel.)
import type { RegisterUploadRequest, UploadConfig, UploadKind, UploadRef } from '../../shared/types';
import { api, ApiClientError } from './api';

export interface UploadProgress { loaded: number; total: number }
export interface UploadOptions {
  filename?: string;
  onProgress?: (p: UploadProgress) => void;
  signal?: AbortSignal;
}

const CSRF_HEADER = { 'X-Requested-With': 'kiit-mailer' };

let configPromise: Promise<UploadConfig> | null = null;

/** GET /api/uploads/config, cached for the page lifetime (reset after a failure). */
export function getUploadConfig(): Promise<UploadConfig> {
  if (!configPromise) {
    configPromise = api.get<UploadConfig>('/uploads/config').catch(err => {
      configPromise = null;
      throw err;
    });
  }
  return configPromise;
}

/** For tests or after an admin changes storage settings. */
export function resetUploadConfigCache() { configPromise = null; }

const EXT_BY_TYPE: Record<string, string> = {
  'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/svg+xml': 'svg', 'image/bmp': 'bmp',
  'application/pdf': 'pdf',
};

/** Safe file name with an extension (blob pathnames and Content-Disposition both use it). */
export function safeFilename(name: string | undefined, contentType: string): string {
  let n = (name || '').split(/[\\/]/).pop() || '';
  n = n.replace(/[^\w.\- ()]+/g, '_').replace(/\s+/g, ' ').trim().slice(0, 120);
  if (!n || n === '.' || n === '..') n = 'file';
  if (!/\.[a-z0-9]{1,8}$/i.test(n)) {
    const ext = EXT_BY_TYPE[contentType];
    if (ext) n += `.${ext}`;
  }
  return n;
}

function abortError(): Error {
  const e = new Error('Upload cancelled');
  e.name = 'AbortError';
  return e;
}

function uploadLocal(file: Blob, kind: UploadKind, filename: string, contentType: string, opts: UploadOptions): Promise<UploadRef> {
  const qs = new URLSearchParams({ kind, filename, contentType });
  return new Promise<UploadRef>((resolve, reject) => {
    if (opts.signal?.aborted) return reject(abortError());
    const xhr = new XMLHttpRequest();
    xhr.open('POST', `/api/uploads/local?${qs.toString()}`);
    xhr.withCredentials = true;
    xhr.setRequestHeader('X-Requested-With', 'kiit-mailer');
    xhr.setRequestHeader('Content-Type', contentType || 'application/octet-stream');
    xhr.responseType = 'text';
    if (opts.onProgress) {
      xhr.upload.onprogress = e => opts.onProgress!({ loaded: e.loaded, total: e.lengthComputable ? e.total : file.size });
    }
    const onAbort = () => xhr.abort();
    opts.signal?.addEventListener('abort', onAbort, { once: true });
    xhr.onabort = () => reject(abortError());
    xhr.onerror = () => reject(new Error('Network error while uploading. Check your connection and try again.'));
    xhr.onload = () => {
      opts.signal?.removeEventListener('abort', onAbort);
      let body: any = null;
      try { body = xhr.responseText ? JSON.parse(xhr.responseText) : null; } catch { /* not JSON */ }
      if (xhr.status >= 200 && xhr.status < 300 && body?.upload) {
        opts.onProgress?.({ loaded: file.size, total: file.size });
        resolve(body.upload as UploadRef);
      } else {
        reject(new ApiClientError(xhr.status, body || { error: `Upload failed (HTTP ${xhr.status})` }));
      }
    };
    xhr.send(file);
  });
}

async function uploadBlob(file: Blob, kind: UploadKind, filename: string, contentType: string, cfg: UploadConfig, opts: UploadOptions): Promise<UploadRef> {
  const { upload } = await import('@vercel/blob/client');
  // In-body images go to the public store (mail clients fetch them); attachments stay private.
  const access = kind === 'image' && cfg.publicImages ? 'public' : 'private';
  const blob = await upload(`${kind === 'image' ? 'images' : 'attachments'}/${filename}`, file, {
    access,
    handleUploadUrl: '/api/uploads/blob-token',
    clientPayload: JSON.stringify({ kind }),
    headers: CSRF_HEADER,
    contentType,
    multipart: file.size > 8 * 1024 * 1024,
    abortSignal: opts.signal,
    onUploadProgress: opts.onProgress ? p => opts.onProgress!({ loaded: p.loaded, total: p.total }) : undefined,
  });
  const req: RegisterUploadRequest = { kind, url: blob.url, filename, contentType, size: file.size };
  const res = await api.post<{ upload: UploadRef }>('/uploads/register', req, { signal: opts.signal });
  return res.upload;
}

export async function uploadFile(file: File | Blob, kind: UploadKind, opts: UploadOptions = {}): Promise<UploadRef> {
  const cfg = await getUploadConfig();
  const contentType = file.type || 'application/octet-stream';
  const filename = safeFilename(opts.filename ?? (file instanceof File ? file.name : undefined), contentType);
  if (kind === 'attachment' && cfg.maxRawAttachmentBytes && file.size > cfg.maxRawAttachmentBytes) {
    const mb = (cfg.maxRawAttachmentBytes / (1024 * 1024)).toFixed(1);
    throw new Error(`"${filename}" is larger than the ${mb} MB attachment limit.`);
  }
  if (cfg.provider === 'blob') return uploadBlob(file, kind, filename, contentType, cfg, opts);
  if (cfg.provider === 'local') return uploadLocal(file, kind, filename, contentType, opts);
  throw new Error(cfg.degradedReason || 'File storage is not configured on this deployment, so files cannot be uploaded.');
}

/** Server-side fetch + rehost of an external image URL (e.g. pasted Google Docs images, which expire). */
export async function rehostImage(url: string): Promise<UploadRef> {
  const res = await api.post<{ upload: UploadRef }>('/uploads/image-from-url', { url });
  return res.upload;
}

export async function deleteUpload(id: string): Promise<void> {
  await api.del<{ ok: true }>(`/uploads/${encodeURIComponent(id)}`);
}

/** Decode a data: URI into a Blob (used for pasted base64 images). */
export function dataUriToBlob(dataUri: string): Blob {
  const m = /^data:([^;,]+)?((?:;[^;,]+)*?)(;base64)?,(.*)$/s.exec(dataUri);
  if (!m) throw new Error('Not a data URI');
  const type = m[1] || 'application/octet-stream';
  const payload = m[4];
  if (m[3]) {
    const bin = atob(payload.replace(/\s+/g, ''));
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new Blob([bytes], { type });
  }
  return new Blob([decodeURIComponent(payload)], { type });
}
