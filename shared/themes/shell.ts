// Email shell: a faithful port of build_email.py's document, header, signature slot and footer.
//
// The shell's own chrome (header, bands, rules, footer) is written with inline styles exactly
// like the skill does, so it survives with or without juice. The body and signature are styled
// by themeCss() through the <style> block, which the render pipeline inlines with juice
// (preserveMediaQueries, removeStyleTags). Client resets live in a `<style data-embed>` block,
// which juice leaves untouched.
import { escapeHtml as esc } from '../text.js';
import { resolveBrand } from './brand.js';
import { mobileCss, msoCss, styleKit, themeCss } from './css.js';
import type { Brand, ShellOptions } from './model.js';
import { getTheme, mix, type ThemeDef } from './tokens.js';

/** Card width in px (the skill used 620; the app standard is 600). */
export const EMAIL_WIDTH = 600;
/** Horizontal body padding inside the card (48px each side). */
export const BODY_PADDING_X = 48;
/** Usable body content width, e.g. for image width attributes. */
export const CONTENT_WIDTH = EMAIL_WIDTH - 2 * BODY_PADDING_X;

const TABLE = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"';

function rule(color: string, h = 1, margin = '0'): string {
  return (
    `${TABLE} style="margin:${margin};"><tr><td height="${h}" bgcolor="${color}" ` +
    `style="height:${h}px;line-height:${h}px;font-size:0;background-color:${color};">&nbsp;</td></tr></table>`
  );
}

/** Centre-weighted rule: the email-safe stand-in for the original template's gradient fade. */
function fadedRule(t: ThemeDef, bg: string, h = 2): string {
  const a = t.colors.accent_rule;
  const segs: [number, string][] = [[15, mix(a, bg, 0.25)], [15, mix(a, bg, 0.6)], [40, a], [15, mix(a, bg, 0.6)], [15, mix(a, bg, 0.25)]];
  const cells = segs
    .map(([w, col]) => `<td width="${w}%" height="${h}" bgcolor="${col}" style="height:${h}px;line-height:${h}px;font-size:0;background-color:${col};">&nbsp;</td>`)
    .join('');
  return `${TABLE}><tr>${cells}</tr></table>`;
}

interface Ctx { t: ThemeDef; brand: Brand; headline: string; sub: string }

function headerDark({ t, brand, headline, sub }: Ctx): string {
  const { colors: c, fonts: f, type: ty } = t;
  const { label } = styleKit(t);
  const divider = mix(c.accent, c.header_bg, 0.35);
  const tagline = brand.tagline
    ? `<p class="e-label e-lk3" style="margin:0;${label(c.eyebrow_on_dark, 12.5, 1.5)}">${esc(brand.tagline)}</p>`
    : '';
  let logoCell = '';
  if (brand.logoUrl) {
    const w = brand.logoWidth;
    logoCell =
      `<td class="e-logocell" width="${w + 24}" valign="middle" style="padding-right:24px;">` +
      `<img class="e-logo" src="${esc(brand.logoUrl)}" alt="${esc(brand.logoAlt)}" width="${w}" ` +
      `style="display:block;width:${w}px;max-width:${w}px;height:auto;border:0;"></td>`;
  }
  const lockupBorder = logoCell ? `border-left:1px solid ${divider};padding-left:24px;` : '';
  let out = `<tr><td bgcolor="${c.header_bg}" style="background-color:${c.header_bg};padding:0;">`;
  out += rule(c.accent_rule, 3);
  out +=
    `${TABLE}><tr><td class="e-pad" style="padding:30px 40px 26px 40px;">${TABLE}><tr>${logoCell}` +
    `<td valign="middle" style="${lockupBorder}">` +
    (brand.eyebrow
      ? `<p class="e-label e-lk1" style="margin:0 0 4px 0;${label(c.eyebrow_on_dark, 11, ty.label_spacing)}">${esc(brand.eyebrow)}</p>`
      : '') +
    (brand.title
      ? `<p class="e-display e-lk2" style="margin:0 0 10px 0;font-family:${f.display};font-size:17px;font-weight:600;line-height:1.3;color:${c.on_dark_soft};">${esc(brand.title)}</p>`
      : '') +
    `<table role="presentation" width="36" cellpadding="0" cellspacing="0" border="0" style="margin:0 0 ${tagline ? 8 : 0}px 0;"><tr>` +
    `<td height="1" bgcolor="${c.accent}" style="height:1px;line-height:1px;font-size:0;background-color:${c.accent};">&nbsp;</td></tr></table>` +
    `${tagline}</td></tr></table></td></tr></table>`;
  if (headline) {
    const bandBorder = mix(c.accent, c.band_bg, 0.2);
    out +=
      `${TABLE} bgcolor="${c.band_bg}" style="background-color:${c.band_bg};border-top:1px solid ${bandBorder};"><tr>` +
      `<td class="e-pad" align="center" style="padding:20px 40px;text-align:center;">` +
      `<h1 class="e-display e-h1" style="margin:0;font-family:${f.display};font-size:${ty.headline_size}px;` +
      `font-weight:${ty.headline_weight};line-height:1.25;letter-spacing:0.3px;color:${c.on_dark};">${esc(headline)}</h1>` +
      (sub ? `<p class="e-label" style="margin:7px 0 0 0;${label(c.eyebrow_on_dark, 12.5, 2)}">${esc(sub)}</p>` : '') +
      '</td></tr></table>';
  }
  out += fadedRule(t, c.header_bg, 2);
  out += '</td></tr>';
  return out;
}

function headerLight({ t, brand, headline, sub }: Ctx): string {
  const { colors: c, fonts: f, type: ty } = t;
  const { label } = styleKit(t);
  let out =
    `<tr><td bgcolor="${c.header_bg}" class="e-pad" align="center" style="background-color:${c.header_bg};` +
    'padding:34px 40px 0 40px;text-align:center;">';
  if (brand.logoUrl) {
    const w = Math.min(brand.logoWidth, 76);
    out +=
      `<img class="e-logo" src="${esc(brand.logoUrl)}" alt="${esc(brand.logoAlt)}" width="${w}" ` +
      `style="display:block;margin:0 auto 14px auto;width:${w}px;max-width:${w}px;height:auto;border:0;">`;
  }
  if (brand.eyebrow) {
    out += `<p class="e-label e-lk1" style="margin:0 0 4px 0;${label(c.accent, 11.5)}">${esc(brand.eyebrow)}</p>`;
  }
  if (brand.title) {
    out +=
      `<p class="e-display e-lk2" style="margin:0 0 ${brand.tagline ? 6 : 18}px 0;font-family:${f.display};font-size:17px;` +
      `line-height:1.35;color:${c.ink};">${esc(brand.title)}</p>`;
  }
  if (brand.tagline) out += `<p class="e-label e-lk3" style="margin:0 0 18px 0;${label(undefined, 11)}">${esc(brand.tagline)}</p>`;
  out += rule(c.accent_rule, 2) + rule(c.header_bg, 2) + rule(c.accent_rule, 1);
  if (headline) {
    out +=
      `<h1 class="e-display e-h1" style="margin:26px 0 0 0;font-family:${f.display};font-size:${ty.headline_size}px;` +
      `font-weight:${ty.headline_weight};line-height:1.3;color:${c.ink};">${esc(headline)}</h1>`;
    if (sub) out += `<p class="e-label" style="margin:8px 0 0 0;${label(undefined, 12)}">${esc(sub)}</p>`;
  }
  out += '</td></tr>';
  return out;
}

/** 'plain': thin navy top rule, small optional logo, text-only lockup, hairline under it. */
function headerPlain({ t, brand, headline, sub }: Ctx): string {
  const { colors: c, fonts: f, type: ty } = t;
  const { label } = styleKit(t);
  let out = `<tr><td bgcolor="${c.card_bg}" style="background-color:${c.card_bg};padding:0;">${rule(c.accent_rule, 3)}</td></tr>`;
  const hasLockup = brand.logoUrl || brand.title || brand.eyebrow;
  if (hasLockup || headline) {
    out += `<tr><td class="e-pad" bgcolor="${c.card_bg}" style="background-color:${c.card_bg};padding:26px 48px 0 48px;">`;
    if (hasLockup) {
      let logoCell = '';
      if (brand.logoUrl) {
        const w = Math.min(brand.logoWidth, 44);
        logoCell =
          `<td width="${w + 16}" valign="middle" style="padding-right:16px;">` +
          `<img src="${esc(brand.logoUrl)}" alt="${esc(brand.logoAlt)}" width="${w}" ` +
          `style="display:block;width:${w}px;max-width:${w}px;height:auto;border:0;"></td>`;
      }
      out +=
        `${TABLE}><tr>${logoCell}<td valign="middle">` +
        (brand.title
          ? `<p class="e-display e-lk2" style="margin:0;font-family:${f.display};font-size:17px;font-weight:700;line-height:1.35;color:${c.accent};">${esc(brand.title)}</p>`
          : '') +
        (brand.eyebrow
          ? `<p class="e-label" style="margin:2px 0 0 0;${label(c.muted, 10.5, 1.2)}line-height:1.5;">${esc(brand.eyebrow)}</p>`
          : '') +
        (brand.tagline
          ? `<p class="e-label" style="margin:2px 0 0 0;${label(c.muted, 10.5, 1.2)}line-height:1.5;">${esc(brand.tagline)}</p>`
          : '') +
        '</td></tr></table>';
      out += rule(c.rule, 1, '20px 0 0 0');
    }
    if (headline) {
      out +=
        `<h1 class="e-display e-h1" style="margin:26px 0 0 0;font-family:${f.display};font-size:${ty.headline_size}px;` +
        `font-weight:${ty.headline_weight};line-height:1.3;color:${c.ink};">${esc(headline)}</h1>`;
      if (sub) out += `<p class="e-label" style="margin:6px 0 0 0;${label(undefined, 11.5)}">${esc(sub)}</p>`;
    }
    out += '</td></tr>';
  }
  return out;
}

function footer(t: ThemeDef, brand: Brand, unsubscribeUrl: string | null): string {
  const c = t.colors;
  const { label, body } = styleKit(t);
  if (!brand.footerLine && !brand.disclaimer && !unsubscribeUrl) return '';
  const align = t.header_style === 'plain' ? 'left' : 'center';
  const line = brand.footerLine
    ? `<p class="e-label" style="margin:0;${label(c.on_dark_muted, 11.5, 1.8)}line-height:1.6;">${esc(brand.footerLine)}</p>`
    : '';
  const disc = brand.disclaimer
    ? `<p class="e-body" style="margin:6px 0 0 0;${body(12, c.on_dark_faint, 1.5)}">${esc(brand.disclaimer)}</p>`
    : '';
  const unsub = unsubscribeUrl
    ? `<p class="e-body" style="margin:8px 0 0 0;${body(12, c.on_dark_faint, 1.5)}">` +
      `<a class="kb-unsubscribe" href="${esc(unsubscribeUrl)}" target="_blank" style="color:${c.on_dark_muted};text-decoration:underline;">Unsubscribe</a></p>`
    : '';
  const inner = `${line}${disc}${unsub}`;
  if (t.header_style === 'light' || t.header_style === 'plain') {
    const ruleColor = t.header_style === 'plain' ? c.rule : c.accent_rule;
    return (
      `<tr><td class="e-pad" bgcolor="${c.card_bg}" style="background-color:${c.card_bg};padding:0 48px 28px 48px;">` +
      `${rule(ruleColor, 1)}${TABLE}><tr><td align="${align}" style="padding:16px 0 0 0;text-align:${align};">${inner}</td></tr></table></td></tr>`
    );
  }
  return (
    `<tr><td bgcolor="${c.header_bg}" style="background-color:${c.header_bg};padding:0;">${fadedRule(t, c.header_bg, 2)}` +
    `${TABLE}><tr><td class="e-pad" align="center" style="padding:16px 40px 18px 40px;text-align:center;">${inner}</td></tr></table></td></tr>`
  );
}

/** Client resets; `data-embed` makes juice leave this block alone (not inlined, not removed). */
const RESET_CSS =
  ':root{color-scheme:light only;supported-color-schemes:light;}' +
  'body,table,td,a{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;}' +
  'table,td{mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;}' +
  'img{-ms-interpolation-mode:bicubic;border:0;outline:none;text-decoration:none;}' +
  'a[x-apple-data-detectors]{color:inherit!important;text-decoration:none!important;}' +
  'u + #body a{color:inherit;text-decoration:none;}';

/** Full HTML document: letterhead header, body in `<td class="kb-body">`, signature, footer. Not yet inlined. */
export function renderShell(opts: ShellOptions): string {
  const t = getTheme(opts.theme);
  const { colors: c, fonts: f } = t;
  const { body } = styleKit(t);
  const brand = resolveBrand(opts.brand);
  const ctx: Ctx = { t, brand, headline: (opts.headline ?? '').trim(), sub: (opts.subheadline ?? '').trim() };
  const header = t.header_style === 'dark' ? headerDark(ctx) : t.header_style === 'light' ? headerLight(ctx) : headerPlain(ctx);
  const sig = opts.signatureHtml && opts.signatureHtml.trim()
    ? `<table role="presentation" class="kb-signature-table" width="100%" cellpadding="0" cellspacing="0" border="0">` +
      `<tr><td class="kb-signature">${opts.signatureHtml}</td></tr></table>`
    : '';
  const fontLink = f.webfont_url
    ? `<!--[if !mso]><!--><link href="${esc(f.webfont_url)}" rel="stylesheet" type="text/css"><!--<![endif]-->\n`
    : '';
  const cardBorder = mix(c.ink, c.page_bg, 0.08);
  const W = EMAIL_WIDTH;

  return `<!DOCTYPE html>
<html lang="en" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="x-apple-disable-message-reformatting">
<meta name="format-detection" content="telephone=no,date=no,address=no,email=no">
<meta name="color-scheme" content="light only">
<meta name="supported-color-schemes" content="light">
<title>${esc(opts.subject)}</title>
<!--[if mso]><noscript><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml></noscript><![endif]-->
${fontLink}<style data-embed>${RESET_CSS}</style>
<style>
${themeCss(opts.theme, '.kb-body')}
${mobileCss(opts.theme, '.kb-body')}
</style>
<!--[if mso]><style>${msoCss(opts.theme, '.kb-body')}</style><![endif]-->
</head>
<body id="body" style="margin:0;padding:0;width:100%;background-color:${c.page_bg};">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="${c.page_bg}" style="background-color:${c.page_bg};">
<tr><td class="e-outer" align="center" style="padding:32px 12px 44px 12px;">
<!--[if mso]><table role="presentation" width="${W}" cellpadding="0" cellspacing="0" border="0" align="center"><tr><td><![endif]-->
<table role="presentation" class="e-card" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="${c.card_bg}" style="max-width:${W}px;width:100%;background-color:${c.card_bg};border:1px solid ${cardBorder};">
${header}
<tr><td class="kb-body e-body e-body-pad" bgcolor="${c.card_bg}" style="background-color:${c.card_bg};padding:38px ${BODY_PADDING_X}px 34px ${BODY_PADDING_X}px;${body()}">
${opts.bodyHtml}
${sig}
</td></tr>
${footer(t, brand, opts.unsubscribeUrl ?? null)}
</table>
<!--[if mso]></td></tr></table><![endif]-->
</td></tr>
</table>
</body>
</html>
`;
}
