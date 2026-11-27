// Pure HTML transforms that turn the server-rendered email into what each client
// would actually show. No DOM access: these run in vitest (node) as well as the browser.
import { applyFontFallbacks } from '../../shared/fonts';
import { GMAIL_CLIP_BYTES } from '../../shared/text';

// ------------------------------------------------------------------ CSS helpers

export interface Decl { prop: string; value: string }

/** Split a declaration list on top-level `;` (ignores `;` inside quotes and parentheses, e.g. data: URLs). */
export function splitDeclarations(css: string): string[] {
  const out: string[] = [];
  let cur = '';
  let depth = 0;
  let quote: string | null = null;
  for (let i = 0; i < css.length; i++) {
    const ch = css[i];
    if (quote) {
      cur += ch;
      if (ch === quote && css[i - 1] !== '\\') quote = null;
      continue;
    }
    if (ch === '"' || ch === "'") { quote = ch; cur += ch; continue; }
    if (ch === '(') depth++;
    else if (ch === ')') depth = Math.max(0, depth - 1);
    if (ch === ';' && depth === 0) { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out.map(s => s.trim()).filter(Boolean);
}

export function parseDeclarations(css: string): Decl[] {
  const decls: Decl[] = [];
  for (const part of splitDeclarations(css)) {
    const i = part.indexOf(':');
    if (i <= 0) continue;
    decls.push({ prop: part.slice(0, i).trim().toLowerCase(), value: part.slice(i + 1).trim() });
  }
  return decls;
}

function serializeDeclarations(decls: Decl[]): string {
  return decls.map(d => `${d.prop}:${d.value}`).join(';');
}

/** Inside a double-quoted style attribute, quotes are encoded as &quot;, whose `;` would break splitting. */
function decodeStyleAttr(css: string): string {
  return css.replace(/&quot;|&#0*34;|&#x0*22;/gi, "'").replace(/&#0*39;|&#x0*27;|&apos;/gi, "'");
}

const STYLE_ATTR_RE = /(\sstyle\s*=\s*)(["'])([\s\S]*?)\2/gi;

/** Rewrite every inline style attribute's declaration list. Returning '' removes the attribute. */
export function mapInlineStyles(html: string, fn: (decls: Decl[]) => Decl[]): string {
  return html.replace(STYLE_ATTR_RE, (_m, pre: string, q: string, css: string) => {
    const next = serializeDeclarations(fn(parseDeclarations(decodeStyleAttr(css))));
    if (!next) return '';
    return `${pre}${q}${q === '"' ? next.replace(/"/g, "'") : next.replace(/'/g, '"')}${q}`;
  });
}

/** Rewrite the contents of every <style> block. */
export function mapStyleBlocks(html: string, fn: (css: string) => string): string {
  return html.replace(/(<style\b[^>]*>)([\s\S]*?)(<\/style>)/gi, (_m, open: string, css: string, close: string) => open + fn(css) + close);
}

/** Remove the named at-rules (with their `{...}` body or trailing `;`) from a stylesheet, brace-aware. */
export function removeAtRules(css: string, names: string[]): string {
  const re = new RegExp(`@(?:${names.join('|')})\\b`, 'gi');
  let out = '';
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(css))) {
    const start = m.index;
    // Find either a `;` (statement at-rule like @import) or the opening `{` of a block.
    let i = start + m[0].length;
    while (i < css.length && css[i] !== '{' && css[i] !== ';') i++;
    let end: number;
    if (i >= css.length) end = css.length;
    else if (css[i] === ';') end = i + 1;
    else {
      let depth = 0;
      end = css.length;
      for (let j = i; j < css.length; j++) {
        if (css[j] === '{') depth++;
        else if (css[j] === '}') { depth--; if (depth === 0) { end = j + 1; break; } }
      }
    }
    out += css.slice(last, start);
    last = end;
    re.lastIndex = end;
  }
  return out + css.slice(last);
}

/** Apply a declaration filter to every `{ ... }` rule body in a stylesheet (non-nested bodies only). */
function mapRuleBodies(css: string, fn: (decls: Decl[]) => Decl[]): string {
  return css.replace(/\{([^{}]*)\}/g, (_m, body: string) => `{${serializeDeclarations(fn(parseDeclarations(body)))}}`);
}

/** Insert markup at the start of <head> (creating one if needed). */
export function injectIntoHead(html: string, markup: string): string {
  if (/<head\b[^>]*>/i.test(html)) return html.replace(/<head\b[^>]*>/i, m => m + markup);
  if (/<html\b[^>]*>/i.test(html)) return html.replace(/<html\b[^>]*>/i, m => `${m}<head>${markup}</head>`);
  return `<head>${markup}</head>${html}`;
}

/** Insert markup at the end of <head> so it wins the cascade over the email's own styles. */
export function appendToHead(html: string, markup: string): string {
  if (/<\/head>/i.test(html)) return html.replace(/<\/head>/i, markup + '</head>');
  return injectIntoHead(html, markup);
}

// ------------------------------------------------------------------ frame preparation

/**
 * Baseline safety for every preview iframe: no scripts (also blocked by the sandbox),
 * links open nowhere instead of navigating the frame, no vertical scrollbar (the frame
 * is auto-sized to its content).
 */
export function prepareFrameHtml(html: string): string {
  const noScripts = html
    .replace(/<script\b[\s\S]*?<\/script\s*>/gi, '')
    .replace(/<script\b[^>]*\/?>/gi, '');
  return injectIntoHead(
    noScripts,
    '<base target="_blank"><style data-kiit-preview>html{overflow-y:hidden!important}</style>',
  );
}

/** Web fonts are not loaded: drop Google Fonts links, @font-face and @import, and use each catalog font's fallback. */
export function stripWebFonts(html: string): string {
  const noLinks = html.replace(/<link\b[^>]*fonts\.(?:googleapis|gstatic)\.com[^>]*>/gi, '');
  return applyFontFallbacks(mapStyleBlocks(noLinks, css => removeAtRules(css, ['font-face', 'import'])));
}

// ------------------------------------------------------------------ Gmail

/** Index into `str` after which the UTF-8 encoding would exceed `maxBytes`. */
function charIndexAtBytes(str: string, maxBytes: number): number {
  let bytes = 0;
  for (let i = 0; i < str.length; i++) {
    const code = str.codePointAt(i)!;
    const len = code < 0x80 ? 1 : code < 0x800 ? 2 : code < 0x10000 ? 3 : 4;
    if (bytes + len > maxBytes) return i;
    bytes += len;
    if (code > 0xffff) i++; // surrogate pair
  }
  return str.length;
}

export function utf8Length(str: string): number {
  let bytes = 0;
  for (let i = 0; i < str.length; i++) {
    const code = str.codePointAt(i)!;
    bytes += code < 0x80 ? 1 : code < 0x800 ? 2 : code < 0x10000 ? 3 : 4;
    if (code > 0xffff) i++;
  }
  return bytes;
}

export interface ClipResult { html: string; clipped: boolean }

/**
 * Gmail shows only the first ~102 KB of the message HTML, then "[Message clipped]".
 * `sizeBytes` is the server's measured size (the decision); the cut is made at the
 * same byte offset in `html`, backed up so it never lands inside a tag. The browser's
 * parser closes any open elements, which is what Gmail effectively shows.
 */
export function gmailClip(html: string, sizeBytes: number = utf8Length(html), limit: number = GMAIL_CLIP_BYTES): ClipResult {
  if (sizeBytes <= limit) return { html, clipped: false };
  let idx = charIndexAtBytes(html, limit);
  const lastOpen = html.lastIndexOf('<', idx - 1);
  const lastClose = html.lastIndexOf('>', idx - 1);
  if (lastOpen > lastClose) idx = lastOpen; // would cut inside a tag
  // Don't leave a half entity (e.g. "&nbs").
  const amp = html.lastIndexOf('&', idx - 1);
  if (amp > -1 && idx - amp < 10 && !html.slice(amp, idx).includes(';')) idx = amp;
  return { html: html.slice(0, idx), clipped: true };
}

/** Gmail (web and apps): web fonts fall back; message clipped above 102 KB. */
export function gmailTransform(html: string, sizeBytes?: number): ClipResult {
  return gmailClip(stripWebFonts(html), sizeBytes ?? utf8Length(html));
}

// ------------------------------------------------------------------ Outlook (Windows) approximation

const OUTLOOK_DROPPED_PROPS = new Set([
  'border-radius', 'border-top-left-radius', 'border-top-right-radius', 'border-bottom-left-radius', 'border-bottom-right-radius',
  '-webkit-border-radius', '-moz-border-radius',
  'box-shadow', '-webkit-box-shadow', 'text-shadow',
  'max-width', 'background-image', 'position',
]);

const URL_OR_GRADIENT_SRC = String.raw`(?:url|(?:repeating-)?(?:linear|radial|conic)-gradient)\((?:[^()]|\([^()]*\))*\)`;
const BG_KEYWORDS = new Set([
  'no', 'repeat', 'center', 'top', 'left', 'right', 'bottom', 'cover', 'contain', 'fixed', 'scroll', 'local',
  'auto', 'none', 'border', 'padding', 'content', 'initial', 'inherit', 'unset', 'space', 'round',
]);

/** The colour component of a `background` shorthand, if any. */
export function backgroundColorOf(value: string): string | null {
  const rest = value.replace(new RegExp(URL_OR_GRADIENT_SRC, 'gi'), ' ');
  const tokens = rest.match(/#[0-9a-f]{3,8}\b|(?:rgb|hsl)a?\([^)]*\)|(?<![-\w])[a-z]+(?![-\w(])/gi) || [];
  for (const t of tokens) {
    if (t.startsWith('#') || /^(?:rgb|hsl)/i.test(t)) return t;
    if (!BG_KEYWORDS.has(t.toLowerCase())) return t;
  }
  return null;
}

function outlookDecls(decls: Decl[]): Decl[] {
  const out: Decl[] = [];
  const hasImage = new RegExp(URL_OR_GRADIENT_SRC, 'i');
  for (const d of decls) {
    if (OUTLOOK_DROPPED_PROPS.has(d.prop)) continue;
    if (d.prop === 'display' && /\b(?:inline-)?(?:flex|grid)\b/i.test(d.value)) continue;
    if (d.prop === 'background' && hasImage.test(d.value)) {
      // Keep only the colour (Outlook honours background colour, not images or gradients).
      const color = backgroundColorOf(d.value);
      if (color) out.push({ prop: 'background-color', value: color });
      continue;
    }
    out.push(d);
  }
  return out;
}

/**
 * Approximate desktop Outlook (Word rendering engine) in a browser:
 * - Outlook-only conditional markup (<!--[if mso]>...) is revealed, unless it is VML,
 *   which browsers can't draw (the regular HTML version stays visible instead);
 * - border-radius, box-shadow, text-shadow, max-width, background images, position and
 *   flex/grid display are removed from inline styles and <style> rules;
 * - @media queries, @font-face and @import are dropped (Outlook ignores them);
 * - web fonts fall back to their universal equivalents.
 * This is an approximation: Word's box model, padding and spacing quirks are not reproduced.
 */
export function outlookTransform(html: string): string {
  let out = html.replace(/<!--\[if ([^\]]*)\]>([\s\S]*?)<!\[endif\]-->/gi, (m, cond: string, inner: string) => {
    if (!/mso/i.test(cond) || /!\s*mso|not\s+mso/i.test(cond)) return m;
    if (/<(?:v|o|w):/i.test(inner)) return m; // VML: leave hidden
    return inner;
  });
  out = mapStyleBlocks(out, css => mapRuleBodies(removeAtRules(css, ['media', 'font-face', 'import', 'supports', 'keyframes']), outlookDecls));
  out = mapInlineStyles(out, outlookDecls);
  out = out.replace(/(<(?:table|td|th|body|tr)\b[^>]*?)\sbackground\s*=\s*(["'])[^"']*\2/gi, '$1');
  return stripWebFonts(out);
}

// ------------------------------------------------------------------ dark mode reference

/**
 * Forced-darkening simulation (the "full colour invert" some apps apply, e.g. Gmail
 * and Outlook mobile apps in dark mode). Images and background images are re-inverted
 * so photos keep their colours. Ported from the v1 preview.
 */
export function darkModeSimulation(html: string): string {
  return appendToHead(
    html,
    `<style data-kiit-dark>
      :root { color-scheme: dark !important; }
      html { background: #fff !important; filter: invert(1) hue-rotate(180deg); }
      img, picture, video, svg, iframe, canvas,
      [style*="background-image"], [background], [data-skip-darken] {
        filter: invert(1) hue-rotate(180deg);
      }
    </style>`,
  );
}

// ------------------------------------------------------------------ per-client pipeline

export type PreviewClientId = 'gmail-desktop' | 'gmail-mobile' | 'apple-iphone' | 'outlook-windows' | 'dark' | 'text';

export interface ClientHtml { html: string; clipped: boolean }

/** The full transform for one preview tab (except plain text). */
export function htmlForClient(client: Exclude<PreviewClientId, 'text'>, html: string, sizeBytes?: number): ClientHtml {
  switch (client) {
    case 'gmail-desktop':
    case 'gmail-mobile': {
      const r = gmailTransform(html, sizeBytes);
      return { html: prepareFrameHtml(r.html), clipped: r.clipped };
    }
    case 'outlook-windows':
      return { html: prepareFrameHtml(outlookTransform(html)), clipped: false };
    case 'dark':
      // Dark re-theming apps are mostly Gmail/Outlook mobile, which also ignore web fonts.
      return { html: prepareFrameHtml(darkModeSimulation(gmailTransform(html, sizeBytes).html)), clipped: false };
    case 'apple-iphone':
    default:
      return { html: prepareFrameHtml(html), clipped: false };
  }
}
