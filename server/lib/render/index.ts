// THE render pipeline. Preview (/api/render/preview), test sends and job sends all
// call renderEmail(), so what the author previews is exactly what is sent.
//
//   1. subject  = substitute(subject, row)
//   2. body     = rich: substitute(bodyHtml, row, escape) → theme shell (+ signature, unsubscribe link)
//                 html: substitute(full document, row, escape), used as-is
//   3. web-font <link> into <head> (Enhanced fonts only)
//   4. preheader (first thing in <body>)
//   5. lockLightMode (color-scheme meta + embedded :root style)
//   6. juice: inline CSS, keep media queries / font-faces / !important
//   text part: bodyText (substituted) or html-to-text of the final HTML.
// Pure (no DB, no network), so it is unit-tested directly.
import juice from 'juice';
import { convert } from 'html-to-text';
import type { ComposeContent, RenderResult, ThemeId } from '../../../shared/types.js';
import { substitute, injectPreheader, lockLightMode } from '../../../shared/text.js';
import { webFontLinkTag, webFontsUsed } from '../../../shared/fonts.js';
import { renderShell, getTheme } from '../../../shared/themes/index.js';
import { lintEmail } from '../lint.js';
import { hostedImageIds } from '../send/images.js';
import type { SenderRow } from '../access.js';

export type RenderSender = Pick<SenderRow, 'signatureHtml'>;

export interface RenderInput {
  content: ComposeContent;
  sender: RenderSender;
  row?: Record<string, string> | null;
  /** Personal one-click URL (or mailto:) shown as the footer "Unsubscribe" link; omit for none. */
  unsubscribeUrl?: string | null;
  recipientsCount?: number;
}

export interface RenderOutput extends RenderResult {
  /** Upload ids of app-hosted images (/api/uploads/:id/content) that the send path turns into CID attachments. */
  hostedImageIds: string[];
}

function insertIntoHead(html: string, fragment: string): string {
  if (!fragment) return html;
  if (/<\/head>/i.test(html)) return html.replace(/<\/head>/i, fragment + '</head>');
  if (/<head[^>]*>/i.test(html)) return html.replace(/<head[^>]*>/i, m => m + fragment);
  if (/<html[^>]*>/i.test(html)) return html.replace(/<html[^>]*>/i, m => m + '<head>' + fragment + '</head>');
  return fragment + html;
}

/** Google Fonts <link> for enhanced fonts used in `fragment` that `doc` does not already load. */
function authorFontLinkTag(fragment: string, doc: string): string {
  const fams = webFontsUsed(fragment)
    .filter(f => f.googleFamily && !doc.includes('family=' + f.googleFamily.split(':')[0]))
    .map(f => `family=${f.googleFamily}`);
  if (!fams.length) return '';
  return `<link href="https://fonts.googleapis.com/css2?${fams.join('&')}&display=swap" rel="stylesheet" type="text/css">`;
}

/** Wrap a bare fragment (HTML mode without <html>/<body>) into a document. */
function ensureDocument(html: string): string {
  if (/<html[\s>]/i.test(html) || /<body[\s>]/i.test(html)) return html;
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head><body>${html}</body></html>`;
}

/** Keep the light-mode lock <style> as a real style tag (juice would inline and drop it). */
function embedLockStyle(html: string): string {
  return html.replace(/<style>(\s*:root\s*\{\s*color-scheme\s*:\s*light only)/i, '<style data-embed>$1');
}

export function htmlToText(html: string): string {
  const headings = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'].map(h => ({ selector: h, options: { uppercase: false } }));
  return convert(html, {
    wordwrap: 78,
    selectors: [
      { selector: 'img', format: 'skip' },
      { selector: 'head', format: 'skip' },
      { selector: 'style', format: 'skip' },
      { selector: 'div[style*="display:none"]', format: 'skip' },
      { selector: 'div[style*="display: none"]', format: 'skip' },
      { selector: 'span[style*="display:none"]', format: 'skip' },
      { selector: 'span[style*="display: none"]', format: 'skip' },
      { selector: 'a', options: { hideLinkHrefIfSameAsText: true, ignoreHref: false, noAnchorUrl: true } },
      ...headings,
      // Email layout tables read as blocks; label/value detail panels stay as two columns.
      { selector: 'table', format: 'block' },
      { selector: 'tr', format: 'block' },
      { selector: 'td', format: 'block', options: { leadingLineBreaks: 1, trailingLineBreaks: 1 } },
      { selector: 'th', format: 'block', options: { leadingLineBreaks: 1, trailingLineBreaks: 1 } },
      { selector: 'table.kb-details', format: 'dataTable', options: { uppercaseHeaderCells: false } },
    ],
  })
    // Spacer cells in the letterhead leave whitespace-only lines (spaces / &nbsp;). Strip
    // trailing whitespace first so runs of blank lines collapse to a single blank line.
    .replace(/[ \t ‌͏]+$/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export function renderEmail(input: RenderInput): RenderOutput {
  const { content, sender } = input;
  const row = input.row || {};
  const subject = substitute(content.subject || '', row).replace(/[\r\n]+/g, ' ').trim();
  // Unknown ids fall back the same way the theme module does (heritage).
  const theme: ThemeId = getTheme(content.theme).id as ThemeId;

  // 2. body
  let html: string;
  let fontLink: string;
  if (content.editorMode === 'html') {
    html = ensureDocument(substitute(content.bodyHtml || '', row, true));
    fontLink = webFontLinkTag(html);
  } else {
    const bodyHtml = substitute(content.bodyHtml || '', row, true);
    const signatureHtml = content.includeSignature && sender.signatureHtml
      ? substitute(sender.signatureHtml, row, true)
      : null;
    html = renderShell({ theme, bodyHtml, signatureHtml, subject, unsubscribeUrl: input.unsubscribeUrl ?? null });
    // The shell already links the theme's own fonts (MSO-guarded); only add fonts the
    // author chose that the theme does not already load.
    fontLink = authorFontLinkTag(bodyHtml + (signatureHtml || ''), html);
  }

  // 3. web fonts
  html = insertIntoHead(html, fontLink);
  // 4. preheader
  html = injectPreheader(html, substitute(content.preheader || '', row));
  // 5. light-mode lock
  html = embedLockStyle(lockLightMode(html));
  // 6. inline CSS
  html = juice(html, {
    preserveMediaQueries: true,
    preserveFontFaces: true,
    applyWidthAttributes: true,
    applyAttributesTableElements: true,
    removeStyleTags: true,
    preserveImportant: true,
  });

  const text = content.bodyText && content.bodyText.trim()
    ? substitute(content.bodyText, row)
    : htmlToText(html);

  return {
    subject,
    html,
    text,
    sizeBytes: Buffer.byteLength(html, 'utf8'),
    webFontsUsed: webFontsUsed(html).map(f => f.label),
    lint: lintEmail({ subject, html, text, recipientsCount: input.recipientsCount }),
    hostedImageIds: hostedImageIds(html),
  };
}
