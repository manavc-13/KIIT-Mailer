// Theme CSS for body content and the custom blocks in shared/blocks.ts.
//
// Rules are deliberately juice-friendly: class, element, descendant, attribute and
// :first-child/:last-child selectors only; plain hex colours; no CSS variables, rgba(),
// gradients, shadows, flex or grid (see the skill's references/theming.md).
//
// Specificity plan (juice resolves per property by specificity, inline styles always win):
//   scope el             (0,1,1)  generic content: p, h1-h3, a, li, table, td ...
//   scope .block-class   (0,2,0)  blocks override every generic property they inherit
//   scope .block[variant] .cell   (0,3,0)  variants override the default block colours
import type { ThemeId } from '../types.js';
import { getTheme, type ThemeDef } from './tokens.js';

type Rule = [selectors: string[], decls: string];

const px = (n: number) => `${n}px`;

function fmt(rules: Rule[], scope: string): string {
  const sel = (s: string) => (s === '&' ? scope : s.startsWith('&') ? scope + s.slice(1) : `${scope} ${s}`);
  return rules.map(([sels, decls]) => `${sels.map(sel).join(', ')}{${decls}}`).join('\n');
}

/** Style fragments shared with the shell (ported from build_email.py's label_css/body_css). */
export function styleKit(t: ThemeDef) {
  const { colors: c, fonts: f, type: ty } = t;
  const body = (size?: number, color?: string, line?: number) =>
    `font-family:${f.body};font-size:${px(size ?? ty.body_size)};line-height:${line ?? ty.body_line};color:${color ?? c.text};`;
  const label = (color?: string, size?: number, spacing?: number) =>
    `font-family:${f.label};font-size:${px(size ?? ty.label_size)};letter-spacing:${px(spacing ?? ty.label_spacing)};` +
    `text-transform:uppercase;font-weight:${ty.label_weight};color:${color ?? c.muted};`;
  return { body, label };
}

/** `{display, body}` font stacks for a theme (every stack ends in a web-safe face). */
export function themeFonts(theme: ThemeId): { display: string; body: string; label: string; ui: string } {
  const f = getTheme(theme).fonts;
  return { display: f.display, body: f.body, label: f.label, ui: f.ui };
}

/** The theme's Google Fonts css2 URL, or '' when it uses no web fonts (slate, plain). */
export function themeWebfontUrl(theme: ThemeId): string {
  return getTheme(theme).fonts.webfont_url;
}

function detailsRules(t: ThemeDef): Rule[] {
  const { colors: c, type: ty } = t;
  const { body, label } = styleKit(t);
  const wrap = 'word-break:normal;overflow-wrap:break-word;';
  if (t.details_style === 'dark') {
    return [
      [['.kb-details'], `margin:0 0 26px 0;width:100%;border-collapse:collapse;background-color:${c.dark_panel_bg};border-radius:${px(ty.radius)};`],
      [['.kb-details-label'], `width:42%;vertical-align:top;padding:6px 12px 6px 24px;border-bottom:0;${body(15, c.on_dark_soft, 1.55)}font-weight:600;`],
      [['.kb-details-value'], `vertical-align:top;padding:6px 24px 6px 0;border-bottom:0;${body(15, c.on_dark_body, 1.55)}${wrap}`],
      [['.kb-details tr:first-child .kb-details-label', '.kb-details tr:first-child .kb-details-value'], 'padding-top:20px;'],
      [['.kb-details tr:last-child .kb-details-label', '.kb-details tr:last-child .kb-details-value'], 'padding-bottom:20px;'],
      [['.kb-details-label p', '.kb-details-value p'], `margin:0;${body(15, c.on_dark_body, 1.55)}`],
      [['.kb-details-label p'], `color:${c.on_dark_soft};font-weight:600;`],
      [['.kb-details-label strong', '.kb-details-value strong'], `color:${c.on_dark_soft};font-weight:600;`],
      [['.kb-details-label a', '.kb-details-value a'], `color:${c.accent};text-decoration:underline;`],
    ];
  }
  const bg = t.details_style === 'wash' ? c.wash_bg : c.panel_bg;
  const tableBorder = t.details_style === 'panel' ? `border:1px solid ${c.panel_border};` : '';
  const labelBorder = t.details_style === 'wash' ? `border-left:3px solid ${c.accent};` : '';
  return [
    [['.kb-details'], `margin:0 0 26px 0;width:100%;border-collapse:collapse;background-color:${bg};${tableBorder}border-radius:${px(ty.radius)};`],
    [['.kb-details-label'], `width:36%;vertical-align:top;padding:8px 14px 4px 24px;border-bottom:0;${labelBorder}` +
      `${label(c.muted, ty.label_size, Math.min(ty.label_spacing, 1.6))}line-height:1.45;`],
    [['.kb-details-value'], `vertical-align:top;padding:5px 24px 5px 0;border-bottom:0;${body(16, c.text, 1.5)}${wrap}`],
    [['.kb-details tr:first-child .kb-details-label'], 'padding-top:28px;'],
    [['.kb-details tr:first-child .kb-details-value'], 'padding-top:25px;'],
    [['.kb-details tr:last-child .kb-details-label'], 'padding-bottom:24px;'],
    [['.kb-details tr:last-child .kb-details-value'], 'padding-bottom:25px;'],
    [['.kb-details-label p'], `margin:0;${label(c.muted, ty.label_size, Math.min(ty.label_spacing, 1.6))}line-height:1.45;`],
    [['.kb-details-value p'], `margin:0;${body(16, c.text, 1.5)}`],
    [['.kb-details-value strong'], `color:${c.ink};font-weight:600;`],
    [['.kb-details-value a'], `color:${c.link};text-decoration:underline;`],
  ];
}

function calloutRules(t: ThemeDef): Rule[] {
  const c = t.colors;
  const { body } = styleKit(t);
  const tone = (prefix: string, bg: string, border: string, fg: string): Rule[] => [
    [[`${prefix}.kb-callout-cell`], `background-color:${bg};border-left:3px solid ${border};color:${fg};`],
    [[`${prefix}.kb-callout-cell p`], `color:${fg};`],
    [[`${prefix}.kb-callout-cell a`, `${prefix}.kb-callout-cell strong`], `color:${fg};`],
  ];
  return [
    [['.kb-callout'], 'margin:0 0 26px 0;width:100%;border-collapse:collapse;'],
    [['.kb-callout-cell'], `padding:14px 18px;border-bottom:0;vertical-align:top;${body(15.5, c.callout_text, 1.6)}` +
      `background-color:${c.callout_bg};border-left:3px solid ${c.callout_border};`],
    [['.kb-callout-cell p'], `margin:0 0 8px 0;${body(15.5, c.callout_text, 1.6)}`],
    [['.kb-callout-cell p:last-child'], 'margin-bottom:0;'],
    [['.kb-callout-cell a'], `color:${c.callout_text};text-decoration:underline;`],
    [['.kb-callout-cell strong'], `color:${c.callout_text};font-weight:600;`],
    [['.kb-callout-cell ul', '.kb-callout-cell ol'], 'margin:0 0 8px 0;'],
    [['.kb-callout-cell li'], `${body(15.5, c.callout_text, 1.6)}`],
    ...tone('.kb-callout[data-variant="warn"] ', c.warn_bg, c.warn_border, c.warn_text),
    ...tone('.kb-callout[data-variant="success"] ', c.success_bg, c.success_border, c.success_text),
  ];
}

function buttonRules(t: ThemeDef): Rule[] {
  const { colors: c, fonts: f, type: ty } = t;
  const r = px(ty.radius);
  const sec = '.kb-button[data-variant="secondary"] ';
  return [
    [['.kb-button'], 'margin:0 0 22px 0;border-collapse:separate;'],
    [['.kb-button-cell'], `padding:0;border-bottom:0;vertical-align:middle;text-align:center;background-color:${c.button_bg};` +
      `border-radius:${r};mso-padding-alt:13px 33px;`],
    [['.kb-button-link'], `display:inline-block;padding:13px 33px;font-family:${f.ui};font-size:16px;font-weight:600;` +
      `letter-spacing:0.4px;line-height:1.2;color:${c.button_text};background-color:${c.button_bg};` +
      `border:1px solid ${c.button_bg};text-decoration:none;border-radius:${r};`],
    [[`${sec}.kb-button-cell`], `background-color:${c.button2_bg};`],
    [[`${sec}.kb-button-link`], `color:${c.button2_text};background-color:${c.button2_bg};border:1px solid ${c.button2_border};`],
  ];
}

function signatureRules(t: ThemeDef): Rule[] {
  const { colors: c, fonts: f } = t;
  const { body } = styleKit(t);
  return [
    [['.kb-signature-table'], 'margin:8px 0 0 0;width:100%;border-collapse:collapse;'],
    [['.kb-signature'], `padding:24px 0 0 0;border-bottom:0;border-top:1px solid ${c.rule};vertical-align:top;${body(14.5, c.text, 1.55)}`],
    [['.kb-signature p'], `margin:0;${body(14.5, c.text, 1.55)}`],
    [['.kb-signature .kb-signature-closing'], `margin:0 0 14px 0;font-size:16px;line-height:1.5;`],
    [['.kb-signature .kb-signature-name'], `font-family:${f.display};font-size:18px;font-weight:600;line-height:1.5;color:${c.ink};`],
    [['.kb-signature .kb-signature-title'], 'font-weight:600;'],
    [['.kb-signature .kb-signature-org'], `color:${c.muted};`],
    [['.kb-signature .kb-signature-contact'], `margin:8px 0 0 0;font-size:13.5px;line-height:1.6;color:${c.ink};`],
    [['.kb-signature strong', '.kb-signature b'], `color:${c.ink};font-weight:600;`],
    [['.kb-signature em', '.kb-signature i'], `color:${c.signature_accent};`],
    [['.kb-signature a'], `color:${c.link};text-decoration:none;`],
    [['.kb-signature .kb-signature-sep'], `color:${c.rule};`],
  ];
}

/**
 * CSS for body content, scoped under `scope` (default '.kb-body'). Used (a) inside the email
 * <style> so juice inlines it and (b) by the editor canvas so writing looks like the sent mail.
 * Styles p, h1-h3, a, ul/ol/li, blockquote, table, hr, img, the custom blocks
 * (.kb-button, .kb-callout, .kb-details, .kb-divider) and the signature (.kb-signature,
 * which the shell places inside the scope). See shared/blocks.ts for markup.
 */
export function themeCss(theme: ThemeId, scope = '.kb-body'): string {
  const t = getTheme(theme);
  const { colors: c, fonts: f, type: ty } = t;
  const { body, label } = styleKit(t);
  const display = (size: number, weight: number | string, line: number) =>
    `font-family:${f.display};font-size:${px(size)};font-weight:${weight};line-height:${line};color:${c.ink};`;

  const rules: Rule[] = [
    [['&'], `${body()}background-color:${c.card_bg};`],
    [['p'], `margin:0 0 18px 0;${body()}`],
    [['h1'], `margin:0 0 16px 0;${display(ty.headline_size, ty.headline_weight, 1.25)}`],
    [['h2'], `margin:30px 0 12px 0;${display(20, 600, 1.3)}`],
    [['h3'], `margin:24px 0 10px 0;${display(17, 600, 1.35)}`],
    [['a'], `color:${c.link};text-decoration:underline;`],
    [['strong', 'b'], 'font-weight:600;'],
    [['ul', 'ol'], 'margin:0 0 18px 0;padding:0 0 0 26px;'],
    [['li'], `margin:0 0 6px 0;${body()}`],
    [['li p'], 'margin:0;'],
    [['li ul', 'li ol'], 'margin:6px 0 0 0;'],
    [['blockquote'], `margin:0 0 18px 0;padding:4px 0 4px 18px;border-left:3px solid ${c.accent};font-style:italic;${body(undefined, c.muted)}`],
    [['blockquote p'], `margin:0 0 8px 0;color:${c.muted};`],
    [['hr'], `border:0;border-top:1px solid ${c.rule};height:0;margin:26px 0;`],
    [['img'], 'border:0;outline:none;text-decoration:none;max-width:100%;height:auto;-ms-interpolation-mode:bicubic;'],
    [['table'], 'border-collapse:collapse;margin:0 0 22px 0;'],
    [['th'], `padding:9px 10px;border-bottom:2px solid ${c.accent};text-align:left;vertical-align:bottom;` +
      `${label(c.muted, 10, 0.8)}line-height:1.4;`],
    [['td'], `padding:9px 10px;border-bottom:1px solid ${c.rule};vertical-align:top;${body(15, c.text, 1.5)}`],
    [['td p', 'th p'], 'margin:0;font-size:15px;line-height:1.5;'],
    ...buttonRules(t),
    ...calloutRules(t),
    ...detailsRules(t),
    [['.kb-divider'], 'margin:6px 0 26px 0;width:100%;border-collapse:collapse;'],
    [['.kb-divider-cell'], `height:1px;line-height:1px;font-size:0;padding:0;border-bottom:0;background-color:${c.divider};mso-line-height-rule:exactly;`],
    ...signatureRules(t),
  ];
  return fmt(rules, scope);
}

/** Mobile rules for the shell (<style>, kept by juice's preserveMediaQueries). Class selectors only. */
export function mobileCss(theme: ThemeId, scope = '.kb-body'): string {
  const t = getTheme(theme);
  const wash = t.details_style === 'wash' ? `border-left:3px solid ${t.colors.accent}!important;` : '';
  return (
    '@media only screen and (max-width:600px){' +
    '.e-outer{padding:0!important;}' +
    '.e-card{border-left:0!important;border-right:0!important;}' +
    '}' +
    '@media only screen and (max-width:480px){' +
    '.e-pad{padding-left:22px!important;padding-right:22px!important;}' +
    '.e-body-pad{padding:30px 22px 28px 22px!important;}' +
    '.e-logocell{width:72px!important;padding-right:14px!important;}' +
    '.e-logo{width:58px!important;max-width:58px!important;}' +
    '.e-lk1{font-size:9.5px!important;letter-spacing:1.2px!important;}' +
    '.e-lk2{font-size:15px!important;}' +
    '.e-h1{font-size:21px!important;}' +
    `${scope} h1{font-size:22px!important;}` +
    `${scope} .kb-details-label,${scope} .kb-details-value{display:block!important;width:auto!important;}` +
    `${scope} .kb-details-label{padding:14px 20px 0 20px!important;}` +
    `${scope} .kb-details-value{padding:2px 20px 12px 20px!important;${wash}}` +
    `${scope} .kb-callout-cell{padding:12px 14px!important;}` +
    '}'
  );
}

/**
 * Outlook (Word renderer) font overrides, placed in an `<!--[if mso]>` style block. Outlook
 * ignores web fonts and can fall back to Times New Roman on a stack it doesn't understand.
 */
export function msoCss(theme: ThemeId, scope = '.kb-body'): string {
  const f = getTheme(theme).fonts;
  const s = (...x: string[]) => x.map(v => `${scope} ${v}`).join(',');
  const darkDetails = getTheme(theme).details_style === 'dark';
  return (
    `.e-display,${s('h1', 'h2', 'h3', '.kb-signature-name')}{font-family:${f.mso_display}!important;}` +
    `.e-body,${scope},${s('p', 'li', 'td', 'blockquote', '.kb-details-value')}` +
    `${darkDetails ? ',' + s('.kb-details-label') : ''}{font-family:${f.mso_body}!important;}` +
    `.e-label,${s('th')}${darkDetails ? '' : ',' + s('.kb-details-label')}{font-family:${f.mso_label}!important;}` +
    `.e-ui,${s('.kb-button-link')}{font-family:${f.mso_ui}!important;}`
  );
}
