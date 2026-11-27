// Email font catalog. Every entry is a complete fallback stack.
//
// universal: system fonts installed on Windows, macOS, iOS and Android (or with
//            a near-identical fallback), so the email looks the same in Gmail,
//            Outlook and Apple Mail.
// enhanced:  web fonts loaded with a Google Fonts <link>. They render in Apple Mail,
//            iOS Mail, Outlook for Mac, Samsung Mail and Thunderbird. Gmail (all
//            platforms) and Outlook for Windows ignore web fonts and use `fallback`.
//            The preview's Gmail/Outlook views show that fallback.

export type FontGroup = 'universal' | 'enhanced';

export interface EmailFont {
  id: string;
  label: string;
  group: FontGroup;
  /** Full CSS font-family value to write into inline styles. */
  stack: string;
  /** For enhanced fonts: what Gmail / Outlook (Windows) will actually show. */
  fallback: string;
  /** For enhanced fonts: Google Fonts css2 family param, e.g. "Playfair+Display:wght@400;700". */
  googleFamily?: string;
  category: 'sans' | 'serif' | 'mono';
}

export const FONT_CATALOG: EmailFont[] = [
  // --- universal
  { id: 'arial', label: 'Arial', group: 'universal', category: 'sans', stack: 'Arial, Helvetica, sans-serif', fallback: 'Arial, Helvetica, sans-serif' },
  { id: 'helvetica', label: 'Helvetica', group: 'universal', category: 'sans', stack: "Helvetica, Arial, sans-serif", fallback: 'Helvetica, Arial, sans-serif' },
  { id: 'verdana', label: 'Verdana', group: 'universal', category: 'sans', stack: 'Verdana, Geneva, sans-serif', fallback: 'Verdana, Geneva, sans-serif' },
  { id: 'tahoma', label: 'Tahoma', group: 'universal', category: 'sans', stack: 'Tahoma, Verdana, Segoe, sans-serif', fallback: 'Tahoma, Verdana, Segoe, sans-serif' },
  { id: 'trebuchet', label: 'Trebuchet MS', group: 'universal', category: 'sans', stack: "'Trebuchet MS', 'Lucida Grande', 'Lucida Sans Unicode', sans-serif", fallback: "'Trebuchet MS', 'Lucida Grande', 'Lucida Sans Unicode', sans-serif" },
  { id: 'system', label: 'System UI', group: 'universal', category: 'sans', stack: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif", fallback: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif" },
  { id: 'calibri', label: 'Calibri', group: 'universal', category: 'sans', stack: "Calibri, Carlito, 'Segoe UI', Arial, sans-serif", fallback: "Calibri, Carlito, 'Segoe UI', Arial, sans-serif" },
  { id: 'georgia', label: 'Georgia', group: 'universal', category: 'serif', stack: "Georgia, 'Times New Roman', Times, serif", fallback: "Georgia, 'Times New Roman', Times, serif" },
  { id: 'times', label: 'Times New Roman', group: 'universal', category: 'serif', stack: "'Times New Roman', Times, serif", fallback: "'Times New Roman', Times, serif" },
  { id: 'palatino', label: 'Palatino', group: 'universal', category: 'serif', stack: "'Palatino Linotype', 'Book Antiqua', Palatino, serif", fallback: "'Palatino Linotype', 'Book Antiqua', Palatino, serif" },
  { id: 'garamond', label: 'Garamond', group: 'universal', category: 'serif', stack: "Garamond, 'Palatino Linotype', 'Book Antiqua', Georgia, serif", fallback: "Garamond, 'Palatino Linotype', 'Book Antiqua', Georgia, serif" },
  { id: 'cambria', label: 'Cambria', group: 'universal', category: 'serif', stack: "Cambria, Caladea, Georgia, serif", fallback: "Cambria, Caladea, Georgia, serif" },
  { id: 'courier', label: 'Courier New', group: 'universal', category: 'mono', stack: "'Courier New', Courier, monospace", fallback: "'Courier New', Courier, monospace" },
  // --- enhanced (web fonts with a matched fallback)
  { id: 'playfair', label: 'Playfair Display', group: 'enhanced', category: 'serif', googleFamily: 'Playfair+Display:ital,wght@0,400;0,600;0,700;1,400', stack: "'Playfair Display', Georgia, 'Times New Roman', serif", fallback: "Georgia, 'Times New Roman', serif" },
  { id: 'eb-garamond', label: 'EB Garamond', group: 'enhanced', category: 'serif', googleFamily: 'EB+Garamond:ital,wght@0,400;0,600;1,400', stack: "'EB Garamond', Garamond, 'Palatino Linotype', Georgia, serif", fallback: "Garamond, 'Palatino Linotype', Georgia, serif" },
  { id: 'merriweather', label: 'Merriweather', group: 'enhanced', category: 'serif', googleFamily: 'Merriweather:ital,wght@0,400;0,700;1,400', stack: "Merriweather, Georgia, serif", fallback: 'Georgia, serif' },
  { id: 'lato', label: 'Lato', group: 'enhanced', category: 'sans', googleFamily: 'Lato:ital,wght@0,400;0,700;1,400', stack: "Lato, 'Helvetica Neue', Arial, sans-serif", fallback: "'Helvetica Neue', Arial, sans-serif" },
  { id: 'open-sans', label: 'Open Sans', group: 'enhanced', category: 'sans', googleFamily: 'Open+Sans:ital,wght@0,400;0,600;0,700;1,400', stack: "'Open Sans', 'Segoe UI', Arial, sans-serif", fallback: "'Segoe UI', Arial, sans-serif" },
  { id: 'roboto', label: 'Roboto', group: 'enhanced', category: 'sans', googleFamily: 'Roboto:ital,wght@0,400;0,500;0,700;1,400', stack: "Roboto, 'Helvetica Neue', Arial, sans-serif", fallback: "'Helvetica Neue', Arial, sans-serif" },
  { id: 'inter', label: 'Inter', group: 'enhanced', category: 'sans', googleFamily: 'Inter:wght@400;500;600;700', stack: "Inter, 'Segoe UI', Helvetica, Arial, sans-serif", fallback: "'Segoe UI', Helvetica, Arial, sans-serif" },
];

export const DEFAULT_FONT_ID = 'arial';

/** Common pixel sizes offered in the editor. */
export const FONT_SIZES_PX = [11, 12, 13, 14, 15, 16, 18, 20, 22, 24, 28, 32, 36];
export const LINE_HEIGHTS = ['1', '1.15', '1.3', '1.5', '1.65', '2'];

const norm = (s: string) => s.toLowerCase().replace(/["']/g, '').split(',')[0].trim();

/** Find a catalog font by its stack or first family name (case/quote-insensitive). */
export function findFont(familyOrStack: string | null | undefined): EmailFont | undefined {
  if (!familyOrStack) return undefined;
  const first = norm(familyOrStack);
  return FONT_CATALOG.find(f => norm(f.stack) === first || f.label.toLowerCase() === first);
}

/**
 * Map an arbitrary font-family (e.g. from a Google Docs paste: "Comfortaa", "Proxima Nova")
 * to the closest catalog font. Unknown families map by generic category.
 */
export function nearestFont(familyOrStack: string | null | undefined): EmailFont {
  const hit = findFont(familyOrStack);
  if (hit) return hit;
  const s = (familyOrStack || '').toLowerCase();
  const byId = (id: string) => FONT_CATALOG.find(f => f.id === id)!;
  if (/mono|courier|consolas|menlo|code/.test(s)) return byId('courier');
  if (/garamond/.test(s)) return byId('garamond');
  if (/playfair|didot|bodoni/.test(s)) return byId('playfair');
  if (/times|serif|georgia|book|baskerville|caslon|cambria|merriweather|lora|crimson|libre|pt serif|spectral|source serif/.test(s) && !/sans/.test(s)) return byId('georgia');
  if (/calibri|carlito/.test(s)) return byId('calibri');
  if (/verdana/.test(s)) return byId('verdana');
  if (/tahoma/.test(s)) return byId('tahoma');
  if (/trebuchet/.test(s)) return byId('trebuchet');
  if (/roboto/.test(s)) return byId('roboto');
  if (/lato/.test(s)) return byId('lato');
  if (/open sans/.test(s)) return byId('open-sans');
  if (/inter\b/.test(s)) return byId('inter');
  return byId('arial');
}

/** Enhanced fonts referenced anywhere in an HTML string (by family name). */
export function webFontsUsed(html: string): EmailFont[] {
  const lower = html.toLowerCase();
  return FONT_CATALOG.filter(f => f.group === 'enhanced' && lower.includes(norm(f.stack)));
}

/** <link> tag loading every enhanced font used in `html` (empty string if none). */
export function webFontLinkTag(html: string): string {
  const fams = webFontsUsed(html).map(f => `family=${f.googleFamily}`);
  if (!fams.length) return '';
  return `<link href="https://fonts.googleapis.com/css2?${fams.join('&')}&display=swap" rel="stylesheet" type="text/css">`;
}

/**
 * Replace every enhanced font stack in `html` with its fallback, i.e. what Gmail and
 * Outlook (Windows) will actually render. Used by the preview's Gmail/Outlook views.
 */
export function applyFontFallbacks(html: string): string {
  let out = html;
  for (const f of FONT_CATALOG) {
    if (f.group !== 'enhanced') continue;
    const fam = norm(f.stack);
    // Matches font-family values whose first family is this web font, with or without quotes.
    const re = new RegExp(`(font-family\\s*:\\s*)(?:&quot;|["'])?${fam.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?:&quot;|["'])?\\s*,[^;"]*`, 'gi');
    out = out.replace(re, (_m, p1) => p1 + f.fallback.replace(/"/g, "'"));
  }
  return out;
}
