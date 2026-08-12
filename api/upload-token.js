// POST /api/upload-token
//
// Issues a short-lived, scoped token that lets the browser upload an
// attachment DIRECTLY to Vercel Blob storage — the file's bytes never touch
// this function, which is what keeps bulk sends under Vercel's 4.5 MB
// request-body cap. Only used when a Blob store is configured
// (BLOB_READ_WRITE_TOKEN present); see api/config.js for the local/inline
// fallbacks used otherwise.
//
// Docs: https://vercel.com/docs/storage/vercel-blob/client-upload
const logger = require('../lib/logger');
const mailer = require('../lib/mailer');

// Guard the require: @vercel/blob is only needed when a Blob store is
// actually configured. Local/.exe runs never hit this route (they use
// api/upload-local.js instead), so don't let a missing package crash the
// whole server on startup.
let handleUpload = null;
try {
    ({ handleUpload } = require('@vercel/blob/client'));
} catch (e) {
    logger.warn('@vercel/blob not installed — /api/upload-token will be unavailable.');
}

module.exports = async (req, res) => {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' });
    }
    if (!process.env.BLOB_READ_WRITE_TOKEN || !handleUpload) {
        return res.status(400).json({ error: 'Blob storage is not configured on this deployment.' });
    }

    try {
        const jsonResponse = await handleUpload({
            body: req.body,
            request: req,
            onBeforeGenerateToken: async () => ({
                access: 'private', // attachments are read back only by our own server (lib/attachments.js), never linked publicly
                allowedContentTypes: undefined, // any type — KIIT staff attach PDFs, images, docs, etc.
                addRandomSuffix: true,
                maximumSizeInBytes: mailer.maxRawAttachmentBytes(),
                // Blobs are cleaned up explicitly via /api/cleanup after each
                // batch, but this is a backstop in case a batch is abandoned
                // mid-flight (browser closed, crash, etc.).
                cacheControlMaxAge: 60 * 60 * 6,
            }),
            onUploadCompleted: async ({ blob }) => {
                logger.info('Blob upload completed: %s', blob.pathname);
            },
        });
        res.status(200).json(jsonResponse);
    } catch (error) {
        logger.error('Upload token error: %s', error.message);
        res.status(400).json({ error: error.message });
    }
};
