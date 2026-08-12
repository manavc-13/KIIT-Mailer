// POST /api/cleanup  { urls?: string[], ids?: string[] }
//
// Deletes uploaded attachments. Blobs are uploaded with access:'private' (see
// api/upload-token.js) so they're never publicly reachable, but IQAC
// documents still shouldn't linger indefinitely once no longer needed — the
// client calls this when an attachment is removed, when attachments are
// cleared, and best-effort on tab close (public/js/attachments.js). It is
// intentionally NOT called right after a send completes, since the same
// attachment may still be needed for a test send or a retry. Best-effort:
// failures are logged, not surfaced to the user.
const fs = require('fs');
const logger = require('../lib/logger');
const { localUploadPath, localMetaPath } = require('../lib/attachments');

module.exports = async (req, res) => {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' });
    }

    const { urls, ids } = req.body || {};
    const results = { deletedUrls: 0, deletedIds: 0, errors: [] };

    if (Array.isArray(urls) && urls.length && process.env.BLOB_READ_WRITE_TOKEN) {
        try {
            const { del } = require('@vercel/blob');
            await del(urls);
            results.deletedUrls = urls.length;
        } catch (err) {
            logger.error('Blob cleanup error: %s', err.message);
            results.errors.push(err.message);
        }
    }

    if (Array.isArray(ids) && ids.length) {
        for (const id of ids) {
            try {
                await fs.promises.unlink(localUploadPath(id));
                results.deletedIds++;
            } catch (err) {
                if (err.code !== 'ENOENT') results.errors.push(err.message);
            }
            try {
                await fs.promises.unlink(localMetaPath(id));
            } catch (err) { /* meta file is optional */ }
        }
    }

    res.status(200).json({ success: true, ...results });
};
