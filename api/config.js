// GET /api/config
// Single source of truth for client-side limits so the frontend never
// hardcodes a size cap that drifts from what the backend actually enforces.
//
// Three attachment providers, chosen automatically:
//   'blob'   - Vercel deployment with a Blob store configured (BLOB_READ_WRITE_TOKEN
//              set). Browser uploads straight to Blob storage; send-mail fetches
//              the URL. Supports attachments up to Gmail's real ~18 MB raw budget.
//   'local'  - Running via `npm start` / the packaged .exe. Browser uploads to a
//              local temp directory; send-mail reads it back from disk. Same
//              ~18 MB budget (only Gmail's cap applies, not Vercel's).
//   'inline' - Deployed on Vercel WITHOUT a Blob store. There is nowhere to
//              stage the file, so it travels base64-encoded inside the same
//              JSON request as the message — capped well under Vercel's
//              4.5 MB body limit. This is a degraded mode; the UI explains why.
const mailer = require('../lib/mailer');

module.exports = async (req, res) => {
    if (req.method !== 'GET') {
        return res.status(405).json({ error: 'Method Not Allowed' });
    }

    const hasBlob = !!process.env.BLOB_READ_WRITE_TOKEN;
    const isVercel = !!process.env.VERCEL;

    let provider, maxRawAttachmentBytes, batchSize, degraded;
    if (hasBlob) {
        provider = 'blob';
        maxRawAttachmentBytes = mailer.maxRawAttachmentBytes();
        batchSize = 10;
        degraded = false;
    } else if (isVercel) {
        provider = 'inline';
        maxRawAttachmentBytes = mailer.INLINE_MAX_RAW_BYTES;
        batchSize = 1;
        degraded = true;
    } else {
        provider = 'local';
        maxRawAttachmentBytes = mailer.maxRawAttachmentBytes();
        batchSize = 10;
        degraded = false;
    }

    res.status(200).json({
        provider,
        degraded,
        degradedReason: degraded
            ? 'No Vercel Blob store is configured for this deployment, so attachments are limited to a smaller size that fits in a single request. Create a Blob store and set BLOB_READ_WRITE_TOKEN to raise this limit to ~18 MB.'
            : null,
        maxRawAttachmentBytes,
        maxEncodedMessageBytes: mailer.GMAIL_MAX_MESSAGE_BYTES,
        softWarnBytes: mailer.SOFT_WARN_BYTES,
        mimeOverheadBytes: mailer.MIME_OVERHEAD_BYTES,
        batchSize,
        gmailDailyLimits: { personal: 500, workspace: 2000 },
    });
};
