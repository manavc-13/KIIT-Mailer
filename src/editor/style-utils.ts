// Small, pure helpers for reading and sanitising inline CSS values.

/** Parse a style attribute into a lowercase-keyed map (last declaration wins). */
export function parseStyle(style: string | null | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  if (!style) return out;
  let buf = '';
  let quote: string | null = null;
  let depth = 0;
  const flush = () => {
    const i = buf.indexOf(':');
    if (i > 0) {
      const k = buf.slice(0, i).trim().toLowerCase();
      const v = buf.slice(i + 1).trim().replace(/\s*!important\s*$/i, '');
      if (k && v) out[k] = v;
    }
    buf = '';
  };
  for (const ch of style) {
    if (quote) { if (ch === quote) quote = null; buf += ch; continue; }
    if (ch === '"' || ch === "'") { quote = ch; buf += ch; continue; }
    if (ch === '(') depth++;
    if (ch === ')') depth = Math.max(0, depth - 1);
    if (ch === ';' && depth === 0) { flush(); continue; }
    buf += ch;
  }
  flush();
  return out;
}

export function styleString(map: Record<string, string | number | null | undefined>): string {
  return Object.entries(map)
    .filter(([, v]) => v !== null && v !== undefined && v !== '')
    .map(([k, v]) => `${k}:${v}`)
    .join(';');
}

/** Convert a CSS length to px (pt, px, in, cm, mm). Returns null for em/%/unknown. */
export function lengthToPx(v: string | null | undefined): number | null {
  if (!v) return null;
  const m = /^(-?\d*\.?\d+)\s*(px|pt|in|cm|mm)?$/i.exec(v.trim());
  if (!m) return null;
  const n = parseFloat(m[1]);
  switch ((m[2] || 'px').toLowerCase()) {
    case 'px': return n;
    case 'pt': return n * (4 / 3);
    case 'in': return n * 96;
    case 'cm': return n * (96 / 2.54);
    case 'mm': return n * (96 / 25.4);
    default: return null;
  }
}

const NAMED_COLORS = new Set([
  'black', 'white', 'red', 'green', 'blue', 'yellow', 'navy', 'maroon', 'gray', 'grey', 'silver', 'teal',
  'olive', 'purple', 'fuchsia', 'aqua', 'lime', 'orange', 'darkblue', 'darkred', 'darkgreen', 'cyan', 'magenta',
  'brown', 'pink', 'gold', 'darkgray', 'darkgrey', 'lightgray', 'lightgrey', 'darkyellow', 'darkcyan', 'darkmagenta',
]);

/** Normalise a colour to a safe CSS value (#hex lower-case, rgb() -> hex). Returns null if unsafe/unknown. */
export function normalizeColor(v: string | null | undefined): string | null {
  if (!v) return null;
  const s = v.trim().toLowerCase();
  if (/^#[0-9a-f]{3}$/.test(s)) return '#' + s.slice(1).split('').map(c => c + c).join('');
  if (/^#[0-9a-f]{6}$/.test(s)) return s;
  if (/^#[0-9a-f]{8}$/.test(s)) return s.slice(0, 7);
  const rgb = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(s);
  if (rgb) {
    if (rgb[4] !== undefined && parseFloat(rgb[4]) === 0) return 'transparent';
    return '#' + [rgb[1], rgb[2], rgb[3]].map(x => Math.min(255, parseInt(x, 10)).toString(16).padStart(2, '0')).join('');
  }
  if (s === 'transparent') return 'transparent';
  if (NAMED_COLORS.has(s)) {
    // Word's "darkyellow" etc. are not CSS; map the few non-standard ones.
    if (s === 'darkyellow') return '#808000';
    return s;
  }
  return null;
}

/** A font-family value safe to embed in a double-quoted style attribute. */
export function safeFontFamily(v: string | null | undefined): string | null {
  if (!v) return null;
  const s = v.replace(/"/g, "'").replace(/[;<>{}\\]/g, '').trim();
  return s || null;
}

/** "16px" / "12pt" / 16 -> "16px" (rounded). Null when invalid. */
export function normalizeFontSize(v: string | number | null | undefined): string | null {
  if (v === null || v === undefined || v === '') return null;
  const px = typeof v === 'number' ? v : lengthToPx(String(v));
  if (px === null || !isFinite(px) || px <= 0) return null;
  return `${Math.max(6, Math.min(96, Math.round(px)))}px`;
}

/** Normalise line-height to a unitless multiplier string ("1.5"). Accepts 1.5, "150%", "1.38". */
export function normalizeLineHeight(v: string | number | null | undefined): string | null {
  if (v === null || v === undefined || v === '') return null;
  const s = String(v).trim();
  let n: number;
  if (/^\d*\.?\d+%$/.test(s)) n = parseFloat(s) / 100;
  else if (/^\d*\.?\d+$/.test(s)) n = parseFloat(s);
  else return null;
  if (!isFinite(n) || n < 0.5 || n > 4) return null;
  return String(Math.round(n * 100) / 100);
}

export function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function escapeAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Link targets allowed in mail: http(s), mailto, tel, or a URL built from placeholders. */
export function safeHref(href: string | null | undefined): string | null {
  if (!href) return null;
  const s = href.trim();
  if (/^(https?:|mailto:|tel:)/i.test(s)) return s;
  if (/^\{[^{}]+\}/.test(s)) return s; // e.g. {RegistrationLink}
  if (s.startsWith('#')) return s;
  return null;
}

/** Image sources allowed in mail: absolute http(s) or our own upload route. Never data: URIs. */
export function safeImageSrc(src: string | null | undefined): string | null {
  if (!src) return null;
  const s = src.trim();
  if (/^https?:\/\//i.test(s)) return s;
  if (s.startsWith('/api/uploads/')) return s;
  return null;
}

export type LinkCheck = { ok: true; href: string; warning?: string } | { ok: false; error: string };

/** Validate a link typed by the author. Accepts bare domains and emails and fixes them up. */
export function checkLink(input: string): LinkCheck {
  let s = input.trim();
  if (!s) return { ok: false, error: 'Enter a link address.' };
  if (/\s/.test(s) && !/^mailto:/i.test(s)) return { ok: false, error: 'A link cannot contain spaces.' };
  if (/^[^@\s/:]+@[^@\s/]+\.[a-z]{2,}$/i.test(s)) s = `mailto:${s}`;
  else if (!/^[a-z][a-z0-9+.-]*:/i.test(s) && !s.startsWith('{') && /^[\w-]+(\.[\w-]+)+(\/|$|\?|#)/.test(s)) s = `https://${s}`;
  if (/^https:\/\/[^/]+\.[^/]+|^https:\/\/localhost/i.test(s) || /^https:\/\/\{/.test(s)) return { ok: true, href: s };
  if (/^mailto:[^@\s]+@[^@\s]+$/i.test(s)) return { ok: true, href: s };
  if (/^tel:\+?[\d\s()-]+$/i.test(s)) return { ok: true, href: s };
  if (/^\{[^{}]+\}/.test(s)) return { ok: true, href: s, warning: 'This link comes from a recipient field; make sure every row has a valid https:// address.' };
  if (/^http:\/\/[^/]+\.[^/]+/i.test(s)) return { ok: true, href: s, warning: 'This link is not secure (http). Mail clients may flag it; use https:// if the site supports it.' };
  return { ok: false, error: 'Use a full https:// address or an email address (mailto:).' };
}
