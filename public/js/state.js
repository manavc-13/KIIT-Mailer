// Shared mutable state + constants. Imported by every module; `state` is a
// single object reference so mutations made in one module are visible to all.

export const STORAGE_SETTINGS = 'kiit_mailer_settings';
export const STORAGE_LOGS = 'kiit_mailer_logs';
export const STORAGE_QUEUE = 'kiit_mailer_pending_queue';
export const STORAGE_TEMPLATES = 'kiit_mailer_templates';
export const STORAGE_DRAFT = 'kiit_mailer_draft';
export const STORAGE_DAILY_COUNT = 'kiit_mailer_daily_count';

export const SPAM_WORDS = [
    'free', 'winner', 'won', 'prize', 'urgent', 'cash', 'click here', 'buy now',
    'limited time', 'act now', 'guarantee', 'risk-free', 'congratulations',
    'lottery', 'discount', 'offer', '!!', '$$', '100%'
];

// Column-role auto-detection aliases (case-insensitive, BOM/whitespace-trimmed
// on the header side before comparison — see recipients.js normalizeHeader()).
export const ROLE_ALIASES = {
    email: ['email', 'e-mail', 'email address', 'mail', 'emailid', 'email id'],
    name: ['name', 'full name', 'recipient', 'recipient name', 'fullname'],
    cc: ['cc', 'cc email', 'copy', 'cc emails', 'cc address'],
    bcc: ['bcc', 'bcc email', 'blind copy', 'bcc emails', 'bcc address'],
};

export const DEFAULT_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light only">
<meta name="supported-color-schemes" content="light">
<title>KIIT Mailer</title>
<style>
  :root { color-scheme: light only; }
  body { margin:0; padding:0; -webkit-text-size-adjust:100%; }
  @media only screen and (max-width:600px) {
    .card { width:100% !important; padding:24px !important; border-radius:0 !important; }
    .btn { width:100% !important; box-sizing:border-box !important; }
  }
</style>
</head>
<body class="body-bg" style="margin:0; padding:0; background-color:#f4f5f7;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="body-bg" style="background-color:#f4f5f7;">
    <tr>
      <td align="center" style="padding:32px 16px;">
        <table role="presentation" class="card" width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px; width:100%; background-color:#ffffff; border-radius:12px; overflow:hidden; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color:#1f2328;">
          <tr>
            <td style="padding:32px 36px 8px 36px;">
              <h1 class="accent" style="margin:0 0 6px 0; font-size:22px; font-weight:700; color:#0b5394; letter-spacing:-0.2px;">
                KIIT University
              </h1>
              <p class="muted" style="margin:0; font-size:14px; color:#6b7280;">
                Bhubaneswar, Odisha &middot; kiit.ac.in
              </p>
            </td>
          </tr>
          <tr><td style="padding:16px 36px 0 36px;"><hr class="divider" style="border:none; border-top:1px solid #e5e7eb; margin:0;"></td></tr>
          <tr>
            <td style="padding:24px 36px;">
              <p style="margin:0 0 14px 0; font-size:16px; line-height:1.6;">Dear {Name},</p>
              <p style="margin:0 0 14px 0; font-size:16px; line-height:1.6;">
                Replace this with your message. You can use placeholders such as
                <strong>{Name}</strong> and <strong>{Email}</strong> — they will be
                substituted per recipient when sending.
              </p>
              <p style="margin:0 0 24px 0; font-size:16px; line-height:1.6;">
                Best regards,<br>
                <strong>KIIT University</strong>
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td bgcolor="#0b5394" class="btn" style="border-radius:6px;">
                    <a href="https://kiit.ac.in" target="_blank" style="display:inline-block; padding:12px 22px; font-size:14px; font-weight:600; color:#ffffff; text-decoration:none;">Visit KIIT</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td class="muted" style="padding:16px 36px 28px 36px; font-size:12px; color:#6b7280; border-top:1px solid #e5e7eb;">
              You are receiving this email from KIIT Mailer.
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

export const DEFAULT_TEXT = `Dear {Name},

Replace this with your plain-text message. Placeholders like {Name} and {Email} will be substituted per recipient.

Best regards,
KIIT University`;

export const state = {
    recipientSource: 'single',  // 'single' | 'csv' | 'manual'
    editorMode: 'html',         // 'html' | 'rich' | 'text'
    pendingEditorMode: null,
    currentStep: 'recipients',  // 'recipients' | 'compose' | 'review'

    singleExtras: [],

    manualColumns: ['Name', 'Email', 'CC', 'BCC'],
    manualRows: [{ Name: '', Email: '', CC: '', BCC: '' }],
    manualRoleMap: { email: 'Email', name: 'Name', cc: 'CC', bcc: 'BCC' },

    csvData: null,
    csvHeaders: [],
    csvRoleMap: { email: null, name: null, cc: null, bcc: null },

    attachments: [], // [{ localName, size, type, status: 'uploading'|'ready'|'error', ref, errorMsg }]

    isSending: false,
    isPaused: false,
    isCancelled: false,
    quill: null,
    credentials: null,
    config: null, // fetched from /api/config

    selectedRows: new Set(),

    sendResults: [],
    sendQueue: [],
    sendCursor: 0,
    sendMeta: null, // { payload, roleMap, attachmentRefs, displayName, replyTo } for the active/resumed batch
    resultsFilter: 'all',

    previewTheme: 'light',
    previewViewport: 'desktop',
    previewRowIndex: 0,
};
