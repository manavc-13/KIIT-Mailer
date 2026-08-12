// POST /api/send-mail
//
// Rewritten from multipart/busboy (one HTTP request PER recipient, attachments
// re-uploaded every time) to a small JSON batch endpoint:
//   - Accepts up to `MAX_BATCH` fully-resolved messages per request, sharing
//     ONE pooled SMTP connection instead of opening a fresh one per email.
//   - Attachments are passed as lightweight references (Blob URL / local temp
//     id / inline base64 — see lib/attachments.js) and resolved into buffers
//     ONCE per request, then reused for every message in the batch. This is
//     what fixes the "payload error under 15 MB" bug: attachment bytes no
//     longer ride inside this request at all (except the small 'inline'
//     fallback, which is capped to fit).
//   - Errors are classified into plain English (lib/mailer.js classifyError)
//     instead of leaking raw SMTP responses.
const logger = require('../lib/logger');
const { createTransporter, classifyError, encodedSize, GMAIL_MAX_MESSAGE_BYTES } = require('../lib/mailer');
const { resolveAttachments } = require('../lib/attachments');

const MAX_BATCH = 10;

module.exports = async (req, res) => {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' });
    }

    const body = req.body || {};
    const { smtpUser, smtpPass, displayName, replyTo, attachmentRefs, messages } = body;

    if (!smtpUser || !smtpPass) {
        return res.status(400).json({ error: 'Missing SMTP Credentials (User/Pass)' });
    }
    if (!Array.isArray(messages) || messages.length === 0) {
        return res.status(400).json({ error: 'No messages provided' });
    }
    if (messages.length > MAX_BATCH) {
        return res.status(400).json({ error: `Too many messages in one batch (max ${MAX_BATCH})` });
    }

    let attachments = [];
    try {
        attachments = await resolveAttachments(attachmentRefs);
    } catch (err) {
        logger.error('Attachment resolution error: %s', err.message);
        return res.status(400).json({ error: `Attachment error: ${err.message}` });
    }

    const attachmentBytes = attachments.reduce((sum, a) => sum + (a.content ? a.content.length : 0), 0);

    let transporter;
    try {
        transporter = createTransporter({ user: smtpUser, pass: smtpPass });
    } catch (err) {
        return res.status(400).json({ error: err.message });
    }

    const results = [];
    for (const msg of messages) {
        const to = (msg.to || '').trim();
        if (!to) {
            results.push({ to: msg.to || '', success: false, error: 'Missing recipient address', retryable: false });
            continue;
        }
        if (!msg.subject) {
            results.push({ to, success: false, error: 'Missing subject', retryable: false });
            continue;
        }
        if (!msg.html && !msg.text) {
            results.push({ to, success: false, error: 'Email body is empty (no html or text provided)', retryable: false });
            continue;
        }

        // Server-side safety net mirroring the client's pre-send guard: total
        // encoded message size must stay under Gmail's real 25 MB cap.
        const bodyBytes = Buffer.byteLength(msg.html || '', 'utf8') + Buffer.byteLength(msg.text || '', 'utf8');
        const estimatedEncoded = encodedSize(attachmentBytes) + bodyBytes;
        if (estimatedEncoded > GMAIL_MAX_MESSAGE_BYTES) {
            results.push({
                to, success: false,
                error: `Message would be ~${(estimatedEncoded / 1024 / 1024).toFixed(1)} MB encoded, over Gmail's 25 MB limit.`,
                retryable: false,
            });
            continue;
        }

        const mailOptions = {
            from: `"${displayName || 'KIIT Mailer'}" <${smtpUser}>`,
            to,
            cc: msg.cc || undefined,
            bcc: msg.bcc || undefined,
            replyTo: replyTo || undefined,
            subject: msg.subject,
            attachments,
        };
        if (msg.html) mailOptions.html = msg.html;
        if (msg.text) mailOptions.text = msg.text;

        try {
            logger.info('Sending email to: %s', to);
            const info = await transporter.sendMail(mailOptions);
            results.push({ to, success: true, messageId: info.messageId });
        } catch (error) {
            const classified = classifyError(error);
            logger.error('Mail error for %s: %s', to, classified.message);
            results.push({ to, success: false, error: classified.message, retryable: classified.retryable });
        }
    }

    transporter.close();
    res.status(200).json({ results });
};
