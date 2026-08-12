// KIIT Mailer - Attachment resolution.
//
// Attachments never travel through the send-mail request body anymore (that's
// what blew past Vercel's 4.5 MB function payload cap). Instead the browser
// uploads them once — to Vercel Blob when deployed, or to a local temp
// directory when running via `npm start` / the packaged .exe — and the send
// handler resolves the small JSON references it's given into real buffers,
// once per batch request, reused across every recipient in that batch.

const fs = require('fs');
const path = require('path');
const os = require('os');

const LOCAL_UPLOAD_DIR = path.join(os.tmpdir(), 'kiit-mailer-uploads');

function ensureLocalUploadDir() {
    if (!fs.existsSync(LOCAL_UPLOAD_DIR)) {
        fs.mkdirSync(LOCAL_UPLOAD_DIR, { recursive: true });
    }
    return LOCAL_UPLOAD_DIR;
}

// A ref looks like one of:
//   { provider: 'blob',   url: 'https://...',    filename, contentType, size }
//   { provider: 'local',  id: 'abc123',           filename, contentType, size }
//   { provider: 'inline', contentBase64: '...',   filename, contentType, size }
// 'inline' is the degraded fallback used only when deployed on Vercel with no
// Blob store configured — see lib/mailer.js INLINE_MAX_RAW_BYTES.
async function resolveOne(ref) {
    if (!ref || (!ref.url && !ref.id && !ref.contentBase64)) {
        throw new Error('Invalid attachment reference');
    }

    if (ref.provider === 'inline' || ref.contentBase64) {
        return {
            filename: ref.filename || 'attachment',
            content: Buffer.from(ref.contentBase64, 'base64'),
            contentType: ref.contentType,
        };
    }

    if (ref.provider === 'local' || ref.id) {
        const dir = ensureLocalUploadDir();
        const safeId = path.basename(String(ref.id)); // defend against path traversal
        const filePath = path.join(dir, safeId);
        if (!fs.existsSync(filePath)) {
            throw new Error(`Attachment "${ref.filename || safeId}" was not found (it may have expired). Please re-attach it.`);
        }
        const content = await fs.promises.readFile(filePath);
        return { filename: ref.filename || safeId, content, contentType: ref.contentType };
    }

    // Vercel Blob reference. Blobs are uploaded with access:'private' (see
    // api/upload-token.js) — a plain fetch() of the URL gets a 400 from Blob
    // storage ("Cannot use public access on a private store"), so this must
    // go through the SDK's get(), which authenticates with BLOB_READ_WRITE_TOKEN.
    const { get } = require('@vercel/blob');
    const result = await get(ref.url, { access: 'private' });
    if (!result) {
        throw new Error(`Attachment "${ref.filename || ref.url}" was not found (it may have expired). Please re-attach it.`);
    }
    const chunks = [];
    for await (const chunk of result.stream) chunks.push(chunk);
    return { filename: ref.filename || 'attachment', content: Buffer.concat(chunks), contentType: ref.contentType };
}

// Resolve a list of refs ONCE and reuse the resulting buffers across every
// message in a batch — this is what stops a bulk send from re-uploading /
// re-downloading the same file N times.
async function resolveAttachments(refs) {
    if (!Array.isArray(refs) || refs.length === 0) return [];
    return Promise.all(refs.map(resolveOne));
}

function localUploadPath(id) {
    return path.join(ensureLocalUploadDir(), path.basename(String(id)));
}

function localMetaPath(id) {
    return path.join(ensureLocalUploadDir(), path.basename(String(id)) + '.meta.json');
}

module.exports = {
    LOCAL_UPLOAD_DIR,
    ensureLocalUploadDir,
    localUploadPath,
    localMetaPath,
    resolveAttachments,
};
