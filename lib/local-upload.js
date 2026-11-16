// Local-dev-only attachment upload, used by the `npm start` dev server
// (local_server.js). Deliberately lives in lib/, not api/: Vercel deploys every
// file under api/ as a serverless function, and the website itself stages
// attachments in Vercel Blob instead (api/upload-token.js).
//
//   POST /api/upload        raw file body, ?filename=&contentType= query params
//                            -> { id, filename, contentType, size }
//   GET  /api/attachment/:id -> streams the file back (for the "view before
//                            send" link in the attachments list)
const fs = require('fs');
const crypto = require('crypto');
const logger = require('./logger');
const { localUploadPath, localMetaPath } = require('./attachments');
const { maxRawAttachmentBytes } = require('./mailer');

async function uploadHandler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' });
    }
    const buf = req.body; // populated by express.raw() on this route
    if (!buf || !Buffer.isBuffer(buf) || buf.length === 0) {
        return res.status(400).json({ error: 'Empty upload body' });
    }
    if (buf.length > maxRawAttachmentBytes()) {
        return res.status(413).json({ error: `File exceeds the ${(maxRawAttachmentBytes() / 1024 / 1024).toFixed(1)} MB attachment limit` });
    }

    const id = crypto.randomUUID();
    const filename = String(req.query.filename || 'attachment');
    const contentType = String(req.query.contentType || 'application/octet-stream');

    try {
        await fs.promises.writeFile(localUploadPath(id), buf);
        await fs.promises.writeFile(localMetaPath(id), JSON.stringify({ filename, contentType, size: buf.length }));
        res.status(200).json({ id, filename, contentType, size: buf.length, provider: 'local' });
    } catch (err) {
        logger.error('Local upload error: %s', err.message);
        res.status(500).json({ error: 'Failed to store upload: ' + err.message });
    }
}

async function serveHandler(req, res) {
    if (req.method !== 'GET') {
        return res.status(405).json({ error: 'Method Not Allowed' });
    }
    const id = req.params.id;
    try {
        const metaRaw = await fs.promises.readFile(localMetaPath(id), 'utf8');
        const meta = JSON.parse(metaRaw);
        const data = await fs.promises.readFile(localUploadPath(id));
        res.setHeader('Content-Type', meta.contentType || 'application/octet-stream');
        res.setHeader('Content-Disposition', `inline; filename="${(meta.filename || 'attachment').replace(/"/g, '')}"`);
        res.status(200).send(data);
    } catch (err) {
        res.status(404).json({ error: 'Attachment not found (it may have expired)' });
    }
}

module.exports = { uploadHandler, serveHandler };
