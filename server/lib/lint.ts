// Pre-send content lint. Pure: runs on the rendered (inlined) HTML, the text
// part and the final subject. Messages are plain English and say what to do.
import type { LintIssue } from '../../shared/types.js';
import { GMAIL_CLIP_BYTES } from '../../shared/text.js';

/**
 * Words and phrases spam filters weigh heavily. Matched case-insensitively on word
 * boundaries. Subject lint uses the full list; body lint skips a few words that are
 * routine in institutional mail body text (e.g. "offer" in an internship offer).
 */
export const SPAM_WORDS: string[] = [
  // money / prizes
  'free', 'winner', 'you have won', 'you won', 'prize', 'cash', 'cash bonus', 'lottery', 'jackpot', 'million dollars',
  'earn money', 'make money', 'extra income', 'double your', 'investment opportunity', 'no fees', 'no cost', 'lowest price',
  'best price', 'cheap', 'discount', 'save big', 'special offer', 'limited offer', 'exclusive deal', 'bonus', 'credit card',
  'loan', 'refund', 'unsecured', 'pre-approved', 'bitcoin', 'crypto',
  // pressure
  'urgent', 'act now', 'act immediately', 'immediately', 'limited time', 'expires today', 'last chance', 'only today',
  "don't miss", 'do not miss', 'hurry', 'once in a lifetime', 'while supplies last', 'apply now', 'order now', 'buy now',
  'call now', 'click here', 'click below', 'click now', 'sign up free', 'get it now',
  // promises
  'guarantee', 'guaranteed', '100% free', '100% satisfied', 'risk-free', 'risk free', 'no obligation', 'no strings attached',
  'miracle', 'amazing', 'incredible deal', 'congratulations', 'you have been selected', 'dear friend', 'winner announced',
  // account scare
  'verify your account', 'account suspended', 'confirm your password', 'update your password', 'unusual activity',
  // legacy list
  'won', 'offer', '100%',
];

/** Words that commonly appear legitimately in KIIT/IQAC body text; checked only in the subject. */
const SUBJECT_ONLY = new Set(['offer', 'won', 'free', 'immediately', 'bonus', 'congratulations', 'amazing', 'apply now', 'refund', 'loan', 'urgent', 'discount', '100%']);

export const URL_SHORTENERS = ['bit.ly', 'tinyurl.com', 'goo.gl', 't.co', 'ow.ly', 'is.gd', 'cutt.ly', 'rb.gy', 'buff.ly', 'shorturl.at', 'tiny.cc', 'rebrand.ly', 'bl.ink', 's.id', 'v.gd'];

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

function findSpamWords(text: string, words: string[]): string[] {
  const hits: string[] = [];
  for (const w of words) {
    const left = /^\w/.test(w) ? '\\b' : '';
    const right = /\w$/.test(w) ? '\\b' : '';
    if (new RegExp(left + escapeRe(w) + right, 'i').test(text)) hits.push(w);
  }
  // Drop hits contained in a longer hit ("free" inside "100% free").
  return hits.filter(h => !hits.some(o => o !== h && o.toLowerCase().includes(h.toLowerCase())));
}

const decodeEntities = (s: string) => s
  .replace(/&nbsp;|&#160;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
  .replace(/&quot;/gi, '"').replace(/&#39;|&apos;/gi, "'").replace(/&#8203;|&zwnj;|&#847;|&#x200c;/gi, '');

/** Strip head/style/script/comments and hidden blocks (preheader), leaving markup for the visible body. */
function visibleBodyHtml(html: string): string {
  let s = html;
  const body = /<body[^>]*>([\s\S]*)<\/body>/i.exec(s);
  if (body) s = body[1];
  return s
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<(style|script|head|title)[^>]*>[\s\S]*?<\/\1>/gi, '')
    // hidden preheader-style blocks: <div style="display:none ...">...</div> (not nested)
    .replace(/<(div|span)[^>]*style=["'][^"']*display\s*:\s*none[^"']*["'][^>]*>[\s\S]*?<\/\1>/gi, '');
}

/** Visible text of an HTML document, whitespace-collapsed. */
export function visibleText(html: string): string {
  return decodeEntities(visibleBodyHtml(html).replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim();
}

function attr(tag: string, name: string): string | null {
  const m = new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, 'i').exec(tag);
  return m ? (m[1] ?? m[2] ?? m[3] ?? '') : null;
}

function hostOf(url: string): string | null {
  try { return new URL(url).hostname.toLowerCase().replace(/^www\./, ''); } catch { return null; }
}

/** If link text looks like a bare URL/domain, its host; otherwise null. */
function urlLikeHost(text: string): string | null {
  const t = text.trim();
  if (!t || /\s/.test(t) || t.includes('@')) return null;
  if (!/^(https?:\/\/)?([a-z0-9-]+\.)+[a-z]{2,}(:\d+)?([/?#]\S*)?$/i.test(t)) return null;
  return hostOf(/^https?:\/\//i.test(t) ? t : 'https://' + t);
}

const sameSite = (a: string, b: string) => a === b || a.endsWith('.' + b) || b.endsWith('.' + a);

/** {Token}-style placeholders still present in visible text/attributes or the subject. */
export function unresolvedPlaceholders(subject: string, html: string): string[] {
  const re = /\{\s*([A-Za-z_][A-Za-z0-9 _.\-]{0,48})\s*\}/g;
  const found = new Set<string>();
  const scan = (s: string) => { for (const m of s.matchAll(re)) found.add(m[1].trim()); };
  scan(subject || '');
  // Visible text plus href/src/alt attributes (all of which reach the reader).
  const body = visibleBodyHtml(html);
  scan(decodeEntities(body.replace(/<[^>]+>/g, ' ')));
  for (const m of body.matchAll(/\s(?:href|src|alt|title)\s*=\s*("[^"]*"|'[^']*')/gi)) scan(decodeURIComponentSafe(m[1]));
  return [...found];
}

function decodeURIComponentSafe(s: string): string { try { return decodeURIComponent(s); } catch { return s; } }

export interface LintInput { subject: string; html: string; text: string | null | undefined; recipientsCount?: number }

export function lintEmail({ subject, html, text }: LintInput): LintIssue[] {
  const issues: LintIssue[] = [];
  const add = (level: LintIssue['level'], code: string, message: string) => issues.push({ level, code, message });
  const subj = (subject || '').trim();

  // ---------------------------------------------------------------- subject
  if (!subj) add('error', 'empty_subject', 'The subject is empty. Add a short, specific subject line.');
  else {
    const letters = subj.replace(/[^A-Za-z]/g, '');
    const upper = letters.replace(/[^A-Z]/g, '').length;
    if (letters.length >= 8 && upper / letters.length > 0.7) {
      add('warn', 'caps_subject', 'The subject is mostly in CAPITAL LETTERS, which spam filters penalise. Use sentence case.');
    }
    if (/!{2,}/.test(subj) || /\${2,}/.test(subj) || /\?{3,}/.test(subj)) {
      add('warn', 'subject_punctuation', 'The subject uses repeated "!", "?" or "$" characters, a common spam signal. Remove them.');
    }
    if (subj.length > 78) {
      add('warn', 'subject_long', `The subject is ${subj.length} characters long. Keep it under 78 so it is not cut off in inboxes.`);
    }
    const hits = findSpamWords(subj, SPAM_WORDS);
    if (hits.length) {
      add('warn', 'spam_word', `The subject contains words spam filters watch for: ${hits.map(h => `"${h}"`).join(', ')}. Consider rephrasing.`);
    }
  }

  // ---------------------------------------------------------------- body
  const bodyHtml = visibleBodyHtml(html || '');
  const bodyText = visibleText(html || '');

  const bodyHits = findSpamWords(bodyText, SPAM_WORDS.filter(w => !SUBJECT_ONLY.has(w)));
  if (bodyHits.length) {
    add('warn', 'spam_word_body', `The message body contains phrases spam filters watch for: ${bodyHits.slice(0, 8).map(h => `"${h}"`).join(', ')}. Consider rephrasing.`);
  }

  // Active content
  if (/<script\b/i.test(html) || /\son[a-z]+\s*=\s*["']/i.test(bodyHtml) || /(href|src)\s*=\s*["']?\s*javascript:/i.test(html)) {
    add('error', 'javascript', 'The email contains JavaScript (a <script> tag, an on… event attribute or a javascript: link). Mail clients block it and spam filters reject it. Remove it.');
  }
  if (/<form\b/i.test(html)) add('error', 'form', 'The email contains a form. Most mail clients disable forms and spam filters penalise them. Link to an online form instead.');
  if (/<(iframe|embed|object)\b/i.test(html)) add('error', 'iframe', 'The email contains an embedded frame or object (<iframe>, <embed> or <object>), which mail clients strip. Link to the content instead.');

  // Images
  const imgs = [...bodyHtml.matchAll(/<img\b[^>]*>/gi)].map(m => m[0]);
  const dataImgs = imgs.filter(t => /^\s*data:/i.test(attr(t, 'src') || '')).length + (/url\(\s*['"]?data:image/i.test(html) ? 1 : 0);
  if (dataImgs) {
    add('error', 'data_uri_image', `${dataImgs} image${dataImgs > 1 ? 's are' : ' is'} embedded as data: URIs. Gmail strips these and spam filters penalise them. Re-insert the image with the image upload button so it is hosted.`);
  }
  const noAlt = imgs.filter(t => attr(t, 'alt') === null).length;
  if (noAlt) {
    add('warn', 'missing_alt', `${noAlt} image${noAlt > 1 ? 's have' : ' has'} no alt text. Add a short description so the email still makes sense when images are blocked.`);
  }
  if (imgs.length) {
    const chars = bodyText.length;
    if (chars < 20) {
      add('warn', 'image_only', 'The email is almost entirely images. Spam filters distrust image-only mail; add the key message as real text.');
    } else if (chars < imgs.length * 100) {
      add('warn', 'low_text_ratio', `There is little text for ${imgs.length} images. Add more written text so the text-to-image balance looks like a normal letter.`);
    }
  }

  // Links
  const anchors = [...bodyHtml.matchAll(/<a\b([^>]*)>([\s\S]*?)<\/a>/gi)];
  const httpLinks = new Set<string>(); const shortLinks = new Set<string>(); const mismatches: string[] = [];
  for (const m of anchors) {
    const href = decodeEntities((attr(' ' + m[1], 'href') || '').trim());
    if (!href || /^(mailto:|tel:|#|cid:)/i.test(href)) continue;
    if (/^http:\/\//i.test(href)) httpLinks.add(href);
    const hh = hostOf(href);
    if (hh && URL_SHORTENERS.some(s => hh === s || hh.endsWith('.' + s))) shortLinks.add(hh);
    const label = decodeEntities(m[2].replace(/<[^>]+>/g, '')).trim();
    const th = urlLikeHost(label);
    if (th && hh && !sameSite(th, hh)) mismatches.push(`"${label}" → ${hh}`);
  }
  if (httpLinks.size) {
    add('warn', 'http_link', `${httpLinks.size} link${httpLinks.size > 1 ? 's use' : ' uses'} insecure http://. Change ${httpLinks.size > 1 ? 'them' : 'it'} to https:// (e.g. ${[...httpLinks][0].slice(0, 80)}).`);
  }
  if (shortLinks.size) {
    add('warn', 'shortener', `Links use URL shorteners (${[...shortLinks].join(', ')}), which spam filters distrust. Use the full destination URL.`);
  }
  if (mismatches.length) {
    add('error', 'link_mismatch', `Link text shows one web address but points to another (${mismatches.slice(0, 3).join('; ')}). This looks like phishing to spam filters. Make the text match the link or use descriptive words instead.`);
  }

  // Unresolved placeholders
  const unresolved = unresolvedPlaceholders(subj, html || '');
  if (unresolved.length) {
    add('error', 'unresolved_placeholder', `No value for ${unresolved.slice(0, 6).map(p => `{${p}}`).join(', ')}. Add a column with ${unresolved.length > 1 ? 'these names' : 'this name'} to the recipient list or remove the placeholder${unresolved.length > 1 ? 's' : ''}.`);
  }

  // Size
  const bytes = Buffer.byteLength(html || '', 'utf8');
  if (bytes > GMAIL_CLIP_BYTES) {
    add('warn', 'gmail_clip', `The email HTML is ${(bytes / 1024).toFixed(0)} KB. Gmail will clip anything over 102 KB and hide the rest behind "View entire message". Shorten the content or remove heavy formatting.`);
  }

  if (!text || !text.trim()) add('info', 'no_text', 'There is no plain-text version. One will be needed for the best deliverability.');

  return issues;
}
