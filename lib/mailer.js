// KIIT Mailer - Transporter factory, size math, and error classification.
//
// Single source of truth for the numbers that matter:
//   - Gmail / Google Workspace caps the total *encoded* message size at 25 MB
//     (https://support.google.com/mail/answer/6584). Base64 (used for MIME
//     attachments) inflates raw bytes by 4/3, so the safe raw-attachment
//     budget is well under 25 MB once headers/body/preheader are accounted for.
//   - Vercel Serverless Functions cap the request body at ~4.5 MB
//     (https://vercel.com/docs/errors/FUNCTION_PAYLOAD_TOO_LARGE). We never
//     put attachment bytes in that request anymore (see lib/attachments.js) —
//     only small JSON — so this ceiling should not be hit in normal use.

const GMAIL_MAX_MESSAGE_BYTES = 25 * 1024 * 1024; // Gmail / Workspace hard cap (encoded)
const SOFT_WARN_BYTES = 10 * 1024 * 1024;          // Above this, many external servers balk
const MIME_OVERHEAD_BYTES = 96 * 1024;             // Headers, boilerplate, preheader, base64 line breaks
const BASE64_RATIO = 4 / 3;

const VERCEL_BODY_LIMIT_BYTES = 4.5 * 1024 * 1024;

// Fallback cap used ONLY when deployed on Vercel with no Blob store configured
// (no BLOB_READ_WRITE_TOKEN). In that mode attachments travel inline, base64,
// inside the same JSON request as the message — so they must fit comfortably
// under Vercel's 4.5 MB body limit alongside the HTML body and JSON overhead.
const INLINE_MAX_RAW_BYTES = 2.8 * 1024 * 1024;

function encodedSize(rawBytes) {
    return Math.ceil((rawBytes / 3)) * 4;
}

// Usable raw attachment budget once base64 inflation + MIME overhead are subtracted.
function maxRawAttachmentBytes() {
    const budgetForEncoded = GMAIL_MAX_MESSAGE_BYTES - MIME_OVERHEAD_BYTES;
    return Math.floor(budgetForEncoded / BASE64_RATIO);
}

function createTransporter({ user, pass }) {
    if (!user || !pass) {
        throw new Error('Missing SMTP Credentials (User/Pass)');
    }
    return require('nodemailer').createTransport({
        service: 'gmail',
        auth: { user, pass },
        pool: true,
        maxConnections: 2,
        maxMessages: 100,
    });
}

// Map raw SMTP/Nodemailer errors to a plain-English message + whether a retry
// is likely to help. Used both in the API response and the client results table.
function classifyError(err) {
    const raw = String((err && (err.response || err.message)) || err || 'Unknown error');
    const code = err && (err.responseCode || err.code);

    const is = (re) => re.test(raw);

    if (code === 535 || code === 534 || is(/535|534|Username and Password not accepted|Application-specific password required/i)) {
        return {
            message: 'App Password rejected — regenerate it at Google Account → Security → App Passwords and update Settings.',
            retryable: false,
        };
    }
    if (is(/552|523|message too large|Message size exceeds/i)) {
        return {
            message: 'Message exceeds Gmail\'s 25 MB limit — remove or shrink attachments.',
            retryable: false,
        };
    }
    if (is(/550-5\.4\.5|452[- ]?4\.2\.1|Daily user sending limit exceeded|Daily sending quota exceeded/i)) {
        return {
            message: 'Daily sending quota exhausted (500/day personal Gmail, 2000/day Workspace). Resume tomorrow.',
            retryable: false,
        };
    }
    if (is(/421|454[- ]?4\.7\.0|4\.7\.0|Temporary System Problem|too many login attempts/i)) {
        return {
            message: 'Gmail is throttling this connection — will retry automatically. Consider increasing the delay.',
            retryable: true,
        };
    }
    if (code === 'ECONNECTION' || code === 'ETIMEDOUT' || code === 'ESOCKET' || is(/ECONNECTION|ETIMEDOUT|ESOCKET|ENOTFOUND|ECONNRESET/i)) {
        return {
            message: 'Network problem reaching smtp.gmail.com — will retry automatically.',
            retryable: true,
        };
    }
    if (is(/FUNCTION_PAYLOAD_TOO_LARGE|413|PayloadTooLarge/i)) {
        return {
            message: 'Request payload too large — this should be unreachable now that attachments upload separately. Please report this.',
            retryable: false,
        };
    }
    return { message: raw, retryable: false };
}

module.exports = {
    GMAIL_MAX_MESSAGE_BYTES,
    SOFT_WARN_BYTES,
    MIME_OVERHEAD_BYTES,
    BASE64_RATIO,
    VERCEL_BODY_LIMIT_BYTES,
    INLINE_MAX_RAW_BYTES,
    encodedSize,
    maxRawAttachmentBytes,
    createTransporter,
    classifyError,
};
