// Email themes / letterheads.
//
// Themes are ported from the iqac-email-html skill (heritage, meridian, letterhead, slate) plus 'plain'.
// Everything here is pure (no DOM, no Node APIs) so it runs in the browser (editor canvas,
// previews) and on the server (render pipeline).
//
//   tokens.ts  colour / font / type tokens per theme (1:1 with the skill's themes/*.json)
//   css.ts     themeCss() for body content + custom blocks + signature, mobile and Outlook CSS
//   shell.ts   renderShell(): the full email document (header bands, card, signature, footer)
//   brand.ts   DEFAULT_BRAND (assets/brand.json) and starterSignatureHtml()
import { THEME_DEFS, THEME_ORDER } from './tokens.js';
import type { ThemeMeta } from './model.js';

export type { ThemeMeta, Brand, ShellOptions, SignatureInput } from './model.js';
export type { ThemeDef, ThemeColors, ThemeFontSet, ThemeType, HeaderStyle, DetailsStyle } from './tokens.js';
export { THEME_DEFS, THEME_ORDER, DEFAULT_THEME, getTheme, mix } from './tokens.js';
export { DEFAULT_BRAND, resolveBrand, starterSignatureHtml, SIGNATURE_DEFAULT_DEPARTMENT, SIGNATURE_ORG_LINE } from './brand.js';
export { themeCss, themeFonts, themeWebfontUrl, mobileCss, msoCss } from './css.js';
export { renderShell, EMAIL_WIDTH, CONTENT_WIDTH, BODY_PADDING_X } from './shell.js';

/** Picker metadata, in display order. */
export const THEMES: ThemeMeta[] = THEME_ORDER.map(id => {
  const t = THEME_DEFS[id];
  return { id, label: t.label, description: t.description, swatch: [...t.swatch] };
});

/** `<link>` tag that loads the theme's web fonts (for the editor canvas / previews), or ''. */
export function themeFontLinkTag(theme: import('../types.js').ThemeId): string {
  const url = THEME_DEFS[theme]?.fonts.webfont_url;
  return url ? `<link href="${url.replace(/&/g, '&amp;')}" rel="stylesheet" type="text/css">` : '';
}
