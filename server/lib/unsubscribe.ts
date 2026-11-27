// Signed per-recipient unsubscribe tokens and the small public pages for /u/:token.
// Pure (no DB): the route does the suppression write.
import { signToken, verifyToken } from './crypto.js';
import { escapeHtml, isValidEmail } from '../../shared/text.js';

export interface UnsubscribePayload { e: string; s: string; j?: string }

/**
 * Public base URL of the app, without a trailing slash. APP_BASE_URL wins; on Vercel
 * without it, the project's production domain; otherwise the local dev server.
 */
export function appBaseUrl(): string {
  const explicit = process.env.APP_BASE_URL?.trim();
  const vercelHost = process.env.VERCEL_PROJECT_PRODUCTION_URL?.trim().replace(/^https?:\/\//i, '');
  const base = explicit || (vercelHost ? `https://${vercelHost}` : 'http://localhost:5173');
  return base.replace(/\/+$/, '');
}

export function makeUnsubscribeToken(email: string, senderAccountId: string, jobId?: string): string {
  return signToken({ e: email.trim().toLowerCase(), s: senderAccountId, ...(jobId ? { j: jobId } : {}) });
}

export function unsubscribeUrl(email: string, senderAccountId: string, jobId?: string): string {
  return `${appBaseUrl()}/u/${makeUnsubscribeToken(email, senderAccountId, jobId)}`;
}

/** A placeholder URL for previews, so the footer link renders without issuing a real token. */
export function previewUnsubscribeUrl(): string {
  return `${appBaseUrl()}/u/preview`;
}

export function parseUnsubscribeToken(token: string): UnsubscribePayload | null {
  if (!token || token.length > 1000) return null;
  let p: UnsubscribePayload | null = null;
  try { p = verifyToken<UnsubscribePayload>(token); } catch { return null; }
  if (!p || typeof p.e !== 'string' || typeof p.s !== 'string' || !isValidEmail(p.e)) return null;
  return p;
}

// ---------------------------------------------------------------- pages
function page(title: string, inner: string): string {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<meta name="color-scheme" content="light">
<title>${escapeHtml(title)} · KIIT IQAC</title>
<style>
  *{box-sizing:border-box}
  body{margin:0;background:#f5f3ee;color:#1d2433;font-family:Georgia,'Times New Roman',serif;-webkit-text-size-adjust:100%}
  .wrap{max-width:560px;margin:0 auto;padding:48px 16px}
  .card{background:#fff;border:1px solid #e2ddd1;border-top:4px solid #1b2a4a;border-radius:6px;padding:32px 28px}
  .eyebrow{font:600 12px/1.4 Arial,Helvetica,sans-serif;letter-spacing:.12em;text-transform:uppercase;color:#9a7b2f;margin:0 0 8px}
  h1{font-size:24px;line-height:1.3;margin:0 0 16px;color:#1b2a4a;font-weight:normal}
  p{font-size:16px;line-height:1.6;margin:0 0 14px}
  .addr{font-family:Arial,Helvetica,sans-serif;font-weight:bold;word-break:break-all}
  button{font:600 15px/1 Arial,Helvetica,sans-serif;background:#1b2a4a;color:#fff;border:0;border-radius:4px;padding:14px 22px;cursor:pointer}
  button:hover{background:#2b3f6b}
  button:focus-visible{outline:3px solid #9a7b2f;outline-offset:2px}
  .muted{font:13px/1.5 Arial,Helvetica,sans-serif;color:#5d6472;margin-top:22px}
</style>
</head>
<body>
<main class="wrap"><div class="card">
<p class="eyebrow">Internal Quality Assurance Cell · KIIT</p>
${inner}
</div>
<p class="muted">Kalinga Institute of Industrial Technology (KIIT), Bhubaneswar, Odisha, India.</p>
</main>
</body>
</html>`;
}

export function confirmPage(email: string): string {
  return page('Unsubscribe', `<h1>Unsubscribe from these emails</h1>
<p>You are about to stop receiving bulk emails sent through the KIIT IQAC mailing service at:</p>
<p class="addr">${escapeHtml(email)}</p>
<form method="post" action="">
  <input type="hidden" name="confirm" value="1">
  <button type="submit">Confirm unsubscribe</button>
</form>
<p class="muted">Individual correspondence addressed to you personally may still reach you. If you did not mean to open this page, simply close it.</p>`);
}

export function donePage(email: string): string {
  return page('Unsubscribed', `<h1>You have been unsubscribed</h1>
<p><span class="addr">${escapeHtml(email)}</span> will no longer receive bulk emails sent through the KIIT IQAC mailing service.</p>
<p>If this was a mistake, please write to the office that sent you the email and ask to be added back.</p>`);
}

export function invalidPage(): string {
  return page('Link not valid', `<h1>This unsubscribe link is not valid</h1>
<p>The link may be incomplete or may have been altered. Please open it again from the original email, or reply to that email and ask to be removed from the list.</p>`);
}

export function previewPage(): string {
  return page('Preview', `<h1>Unsubscribe link preview</h1>
<p>This is a preview link. In a sent email each recipient gets a personal unsubscribe link here.</p>`);
}
