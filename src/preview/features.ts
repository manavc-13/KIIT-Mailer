// Detects the HTML/CSS features an email uses, keyed by caniemail feature slug.
// Pure string analysis (no DOM), so it runs in tests and in the browser alike.
import { webFontsUsed } from '../../shared/fonts';
import type { FeatureSlug } from './caniemail-spec';
import { parseDeclarations, type Decl } from './transforms';

export interface DetectedFeature {
  slug: FeatureSlug;
  /** Short human hint of where it was seen, e.g. "on 3 elements" or "in <style>". */
  evidence: string;
}

interface StyledDecl extends Decl { tag: string | null; source: 'inline' | 'block' }

/** Remove the hidden preheader block so its deliberate display:none / max-width:0 tricks don't count. */
function stripPreheader(html: string): string {
  return html.replace(/<div\b[^>]*mso-hide\s*:\s*all[^>]*>[\s\S]*?<\/div>/gi, '');
}

function collect(html: string): { decls: StyledDecl[]; blockCss: string; inlineCss: string } {
  const decls: StyledDecl[] = [];
  let inlineCss = '';
  for (const m of html.matchAll(/<([a-z][\w:-]*)\b[^>]*?\sstyle\s*=\s*(["'])([\s\S]*?)\2/gi)) {
    const css = m[3].replace(/&quot;|&#0*34;/gi, "'");
    inlineCss += css + ';';
    for (const d of parseDeclarations(css)) decls.push({ ...d, tag: m[1].toLowerCase(), source: 'inline' });
  }
  let blockCss = '';
  for (const m of html.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)) blockCss += m[1] + '\n';
  // Rule bodies inside <style> (including those nested in @media).
  for (const m of blockCss.matchAll(/\{([^{}]*)\}/g)) {
    for (const d of parseDeclarations(m[1])) decls.push({ ...d, tag: null, source: 'block' });
  }
  return { decls, blockCss, inlineCss };
}

const plural = (n: number, one: string, many = one + 's') => `${n} ${n === 1 ? one : many}`;

function where(list: StyledDecl[]): string {
  const inline = list.filter(d => d.source === 'inline').length;
  const block = list.length - inline;
  const parts: string[] = [];
  if (inline) parts.push(`inline on ${plural(inline, 'element')}`);
  if (block) parts.push(`in <style> (${plural(block, 'rule')})`);
  return parts.join(', ');
}

const ZERO = /^0(?:px|em|rem|%)?$/i;

/** Every caniemail feature the email uses, in a stable order. */
export function detectFeatures(rawHtml: string): DetectedFeature[] {
  const html = stripPreheader(rawHtml);
  const { decls, blockCss } = collect(html);
  const found = new Map<FeatureSlug, string>();
  const add = (slug: FeatureSlug, evidence: string) => { if (!found.has(slug)) found.set(slug, evidence); };
  const byProp = (test: (d: StyledDecl) => boolean, slug: FeatureSlug) => {
    const hits = decls.filter(test);
    if (hits.length) add(slug, where(hits));
  };
  const prop = (re: RegExp) => (d: StyledDecl) => re.test(d.prop);
  const nonZero = (d: StyledDecl) => !ZERO.test(d.value.replace(/!important/i, '').trim());
  const value = (re: RegExp) => (d: StyledDecl) => re.test(d.value);
  const count = (re: RegExp) => (html.match(re) || []).length;

  // --- <style> and at-rules
  const styleBlocks = count(/<style\b[^>]*>/gi);
  if (styleBlocks) add('html-style', plural(styleBlocks, '<style> block'));
  const media = count(/@media\b/gi);
  if (media) add('css-at-media', plural(media, 'media query', 'media queries'));
  if (/prefers-color-scheme/i.test(blockCss)) add('css-at-media-prefers-color-scheme', 'dark-mode styles in <style>');
  const fonts = webFontsUsed(html);
  if (/@font-face/i.test(blockCss) || /<link\b[^>]*fonts\.googleapis\.com/i.test(html) || fonts.length) {
    add('css-at-font-face', fonts.length ? `web fonts: ${fonts.map(f => f.label).join(', ')}` : 'web font declared');
  }
  if (/@import\b/i.test(blockCss)) add('css-at-import', 'in <style>');
  if (/@keyframes\b/i.test(blockCss)) add('css-at-keyframes', 'in <style>');
  if (/:hover\b/i.test(blockCss)) add('css-pseudo-class-hover', 'hover styles in <style>');

  // --- properties
  byProp(d => /^--/.test(d.prop) || /var\(\s*--/.test(d.value), 'css-variables');
  byProp(d => /^(?:-webkit-|-moz-)?border(?:-(?:top|bottom)-(?:left|right))?-radius$/.test(d.prop) && nonZero(d), 'css-border-radius');
  byProp(d => /^(?:-webkit-)?box-shadow$/.test(d.prop) && !/^none/i.test(d.value), 'css-box-shadow');
  byProp(d => d.prop === 'text-shadow' && !/^none/i.test(d.value), 'css-text-shadow');
  byProp(d => (d.prop === 'background-image' && !/^none/i.test(d.value)) || (d.prop === 'background' && /url\(/i.test(d.value)), 'css-background-image');
  byProp(prop(/^background-size$/), 'css-background-size');
  byProp(value(/linear-gradient\(/i), 'css-linear-gradient');
  byProp(value(/radial-gradient\(/i), 'css-radial-gradient');
  byProp(d => d.prop === 'max-width' && !/^none/i.test(d.value) && nonZero(d), 'css-max-width');
  byProp(d => d.prop === 'min-width' && nonZero(d), 'css-min-width');
  byProp(d => d.prop === 'max-height' && !/^none/i.test(d.value) && nonZero(d), 'css-max-height');
  byProp(d => d.prop === 'display' && /\b(?:inline-)?flex\b/i.test(d.value), 'css-display-flex');
  byProp(d => d.prop === 'display' && /\b(?:inline-)?grid\b/i.test(d.value), 'css-display-grid');
  byProp(d => d.prop === 'display' && /^none/i.test(d.value), 'css-display-none');
  byProp(d => /^(?:row-|column-)?gap$/.test(d.prop) && nonZero(d), 'css-gap');
  byProp(d => d.prop === 'position' && !/^static/i.test(d.value), 'css-position');
  byProp(prop(/^z-index$/), 'css-z-index');
  byProp(d => d.prop === 'float' && !/^none/i.test(d.value), 'css-float');
  // Margin is everywhere; flag only the cases clients actually break: on images, or negative.
  {
    const hits = decls.filter(d => /^margin(?:-(?:top|right|bottom|left))?$/.test(d.prop) && (d.tag === 'img' || /(?:^|\s)-\d/.test(d.value)));
    if (hits.length) {
      const onImg = hits.filter(d => d.tag === 'img').length;
      add('css-margin', onImg ? `on ${plural(onImg, '<img>')}` : 'negative margin');
    }
  }
  // Padding only works on table cells in desktop Outlook.
  {
    const hits = decls.filter(d => d.source === 'inline' && /^padding/.test(d.prop) && nonZero(d) && d.tag && !['td', 'th'].includes(d.tag));
    if (hits.length) {
      const tags = [...new Set(hits.map(d => `<${d.tag}>`))].slice(0, 3).join(', ');
      add('css-padding', `on non-table-cell elements (${tags})`);
    }
  }
  byProp(prop(/^opacity$/), 'css-opacity');
  byProp(d => /^overflow(?:-[xy])?$/.test(d.prop) && !/^visible/i.test(d.value), 'css-overflow');
  byProp(prop(/^object-fit$/), 'css-object-fit');
  byProp(d => /^(?:-webkit-)?transform$/.test(d.prop) && !/^none/i.test(d.value), 'css-transform');
  byProp(prop(/^transition/), 'css-transition');
  byProp(d => /^animation/.test(d.prop) && !/^none/i.test(d.value), 'css-animation');
  byProp(value(/\d(?:\.\d+)?rem\b/i), 'css-unit-rem');
  byProp(value(/\d(?:\.\d+)?(?:vw|vh|vmin|vmax)\b/i), 'css-unit-vw');
  byProp(value(/calc\(/i), 'css-unit-calc');
  byProp(value(/clamp\(/i), 'css-function-clamp');
  byProp(d => /^font(?:-family)?$/.test(d.prop) && /(?:^|[\s,'"])system-ui\b/i.test(d.value), 'css-sytem-ui');

  // --- HTML
  const bgAttr = count(/<(?:table|td|th|tr|body)\b[^>]*\sbackground\s*=/gi);
  if (bgAttr) add('html-background', `background="" on ${plural(bgAttr, 'element')}`);
  const svg = count(/<svg[\s>]/gi);
  if (svg) add('html-svg', plural(svg, 'inline <svg>'));
  const imgSrcs = [...html.matchAll(/<img\b[^>]*?\ssrc\s*=\s*(["'])([^"']*)\1/gi)].map(m => m[2].trim());
  const imgCount = (re: RegExp) => imgSrcs.filter(s => re.test(s)).length;
  const svgImg = imgCount(/\.svg(?:[?#]|$)/i);
  if (svgImg) add('image-svg', plural(svgImg, 'SVG image'));
  const video = count(/<video[\s>]/gi);
  if (video) add('html-video', plural(video, '<video>'));
  const audio = count(/<audio[\s>]/gi);
  if (audio) add('html-audio', plural(audio, '<audio>'));
  const picture = count(/<picture[\s>]/gi);
  if (picture) add('html-picture', plural(picture, '<picture>'));
  const srcset = count(/\ssrcset\s*=/gi);
  if (srcset) add('html-srcset', plural(srcset, 'srcset attribute'));
  const form = count(/<(?:form|input|select|textarea)[\s>]/gi);
  if (form) add('html-form', plural(form, 'form element'));
  const sem = [...new Set((html.match(/<(?:header|footer|section|article|nav|main|aside)(?=[\s>])/gi) || []).map(t => t.toLowerCase() + '>'))];
  if (sem.length) add('html-semantics', sem.join(', '));
  const webp = imgCount(/\.webp(?:[?#]|$)/i);
  if (webp) add('image-webp', plural(webp, 'WebP image'));
  const avif = imgCount(/\.avif(?:[?#]|$)/i);
  if (avif) add('image-avif', plural(avif, 'AVIF image'));
  const gif = imgCount(/\.gif(?:[?#]|$)/i);
  if (gif) add('image-gif', plural(gif, 'GIF image'));
  const b64 = imgCount(/^data:image\//i);
  if (b64) add('image-base64', plural(b64, 'embedded data: image'));

  return [...found].map(([slug, evidence]) => ({ slug, evidence }));
}

/** True when the HTML carries Outlook-only conditional markup (ghost tables, VML). */
export function hasOutlookConditionals(html: string): boolean {
  return /<!--\[if\s[^\]]*mso/i.test(html);
}
