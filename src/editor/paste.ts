// Google Docs / Word / web clipboard normalizer.
//
// Used by the editor's `transformPastedHTML` (mode 'paste') and when loading legacy or
// HTML-mode content into the rich editor (mode 'load'). Produces clean HTML that the
// TipTap schema parses predictably:
//  - span styles -> <strong>/<em>/<u>/<s>/<sup>/<sub> + one <span style> carrying only
//    colour, highlight, catalog font stack and px font size
//  - paragraph styles -> text-align, unitless line-height, px margins, data-indent levels
//  - Docs flat/aria-level lists and Word mso-list paragraphs -> properly nested ul/ol
//  - tables keep cell backgrounds and column widths
//  - google.com/url?q= redirects unwrapped; scripts, styles, comments, mso junk removed
//  - images marked data-kb-rehost="1" when they must be re-uploaded (data: URIs always,
//    external URLs on paste) so the editor can swap in hosted URLs asynchronously.
// Requires a DOM (browser, or jsdom in tests).
import { nearestFont } from '../../shared/fonts';
import { MAX_IMAGE_WIDTH, MAX_INDENT, INDENT_PX, fieldRegex } from './constants';
import {
  escapeHtml, lengthToPx, normalizeColor, normalizeLineHeight, parseStyle, safeHref, styleString,
} from './style-utils';

export type PasteSource = 'docs' | 'word' | 'html';
export interface NormalizeOptions {
  /** 'paste' (default): clipboard content. 'load': existing template HTML (no external image rehosting, keeps blank paragraphs). */
  mode?: 'paste' | 'load';
}

/** One Docs/Word indent step is 36pt = 48px. */
const SOURCE_INDENT_PX = 48;
const MAX_SPACING_PX = 48;
const BLACKS = new Set(['#000000', 'black', 'windowtext']);
const DEFAULT_LINK_COLORS = new Set(['#1155cc', '#0563c1', '#467886', '#0000ff', 'blue', '#954f72']);
const DROP_BG = new Set(['transparent', '#ffffff', 'white', 'none']);
const DOCS_DEFAULT_FONTS = new Set(['arial']);
const WORD_DEFAULT_FONTS = new Set(['calibri', 'calibri light', 'aptos', 'aptos display', 'times new roman']);
const KB_CLASS = /^kb-[a-z-]+$/;
const KEEP_ATTRS = new Set([
  'style', 'href', 'src', 'alt', 'width', 'colspan', 'rowspan', 'colwidth', 'start', 'type', 'align',
  'data-field', 'data-indent', 'data-kb-rehost', 'data-align', 'data-variant', 'class',
]);
const BLOCK_TAGS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI', 'TABLE', 'DIV', 'BLOCKQUOTE', 'HR', 'PRE', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'MAIN', 'CENTER', 'ADDRESS', 'FIGURE', 'IMG']);
const CONTAINER_TAGS = new Set(['SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'MAIN', 'CENTER', 'ADDRESS', 'FIGURE', 'FIGCAPTION', 'NAV', 'ASIDE', 'FORM', 'BODY', 'HTML']);

export function detectSource(html: string): PasteSource {
  if (/id=["']?docs-internal-guid/i.test(html)) return 'docs';
  if (/urn:schemas-microsoft-com:office|class=["']?Mso|mso-[a-z-]+\s*:|<o:p>/i.test(html)) return 'word';
  return 'html';
}

/** Unwrap https://www.google.com/url?q=<real>&sa=D... redirects added by Google Docs. */
export function unwrapGoogleRedirect(href: string): string {
  try {
    const u = new URL(href);
    if (/(^|\.)google\.[a-z.]+$/i.test(u.hostname) && u.pathname === '/url') {
      const q = u.searchParams.get('q') || u.searchParams.get('url');
      if (q) return q;
    }
  } catch { /* not absolute */ }
  return href;
}

/** True when the image is already on our storage (no rehost needed). */
export function isHostedImage(src: string): boolean {
  if (src.startsWith('/api/uploads/')) return true;
  try {
    const u = new URL(src);
    if (/\.blob\.vercel-storage\.com$/i.test(u.hostname)) return true;
    if (typeof location !== 'undefined' && u.origin === location.origin && u.pathname.startsWith('/api/uploads/')) return true;
  } catch { /* ignore */ }
  return false;
}

// ------------------------------------------------------------------ string pre-clean

function preClean(html: string): string {
  let s = html;
  const frag = /<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/i.exec(s);
  if (frag) s = frag[1];
  // Word list markers live in downlevel-revealed conditionals; keep them in a marker span
  // so the list converter can read "1." / "a)" / "·" before they are removed.
  s = s.replace(/<!\[if !supportLists\]>([\s\S]*?)<!\[endif\]>/gi, '<span data-kb-listmarker="1">$1</span>');
  s = s.replace(/<!--\[if !supportLists\]-->([\s\S]*?)<!--\[endif\]-->/gi, '<span data-kb-listmarker="1">$1</span>');
  // Conditional comments (VML, mso-only blocks) and their content.
  s = s.replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, '');
  // Downlevel-revealed markers: keep their content (e.g. the non-VML <img>).
  s = s.replace(/<!\[(?:if [^\]]*|endif)\]>/gi, '');
  s = s.replace(/<xml[\s\S]*?<\/xml>/gi, '');
  // Office namespaced tags (<o:p>, <v:shape>, <w:...>, <st1:...>): drop the tag, keep text.
  s = s.replace(/<\/?[a-z]+:[a-z][\w-]*[^>]*>/gi, '');
  return s;
}

// ------------------------------------------------------------------ helpers

function unwrap(el: Element) {
  const parent = el.parentNode;
  if (!parent) return;
  while (el.firstChild) parent.insertBefore(el.firstChild, el);
  parent.removeChild(el);
}

function rename(el: Element, tag: string): HTMLElement {
  const doc = el.ownerDocument;
  const n = doc.createElement(tag);
  for (const a of Array.from(el.attributes)) n.setAttribute(a.name, a.value);
  while (el.firstChild) n.appendChild(el.firstChild);
  el.replaceWith(n);
  return n;
}

/** Space-separated kb-* classes of an element ('' if none). */
export function kbClasses(el: Element): string {
  return (el.getAttribute('class') || '').split(/\s+/).filter(c => KB_CLASS.test(c)).join(' ');
}

const firstFamily = (v: string) => v.split(',')[0].replace(/["']/g, '').trim().toLowerCase();

function hasBlockChild(el: Element): boolean {
  return Array.from(el.children).some(c => BLOCK_TAGS.has(c.tagName));
}

function isBlankText(el: Element): boolean {
  return !(el.textContent || '').replace(/[\s\u00a0]/g, '') && !el.querySelector('img,br,hr,table,[data-field]');
}

function marginParts(st: Record<string, string>): { top?: string; bottom?: string; left?: string } {
  const out: { top?: string; bottom?: string; left?: string } = {};
  if (st.margin) {
    const p = st.margin.split(/\s+/);
    const [t, r = t, b = t, l = r] = p;
    void r;
    out.top = t; out.bottom = b; out.left = l;
  }
  if (st['margin-top']) out.top = st['margin-top'];
  if (st['margin-bottom']) out.bottom = st['margin-bottom'];
  if (st['margin-left']) out.left = st['margin-left'];
  return out;
}

// ------------------------------------------------------------------ lists

interface ListItemSpec { level: number; ordered: boolean; olType: string | null; li: HTMLElement; start?: number }

function olTypeFromStyle(v: string | null | undefined): string | null {
  const s = (v || '').trim();
  if (/^[1aAiI]$/.test(s)) return s;
  switch (s.toLowerCase()) {
    case 'decimal': return '1';
    case 'lower-alpha': case 'lower-latin': return 'a';
    case 'upper-alpha': case 'upper-latin': return 'A';
    case 'lower-roman': return 'i';
    case 'upper-roman': return 'I';
    default: return null;
  }
}

function buildLists(doc: Document, items: ListItemSpec[]): DocumentFragment {
  const frag = doc.createDocumentFragment();
  const stack: { el: HTMLElement; level: number; ordered: boolean }[] = [];
  const top = () => stack[stack.length - 1];
  const make = (it: ListItemSpec) => {
    const el = doc.createElement(it.ordered ? 'ol' : 'ul');
    if (it.ordered && it.olType && it.olType !== '1') el.setAttribute('type', it.olType);
    if (it.ordered && it.start && it.start > 1) el.setAttribute('start', String(it.start));
    return el;
  };
  for (const it of items) {
    const L = Math.max(1, Math.min(9, it.level || 1));
    while (stack.length && top().level > L) stack.pop();
    if (stack.length && top().level === L && top().ordered !== it.ordered) stack.pop();
    if (!stack.length) {
      const el = make(it);
      frag.appendChild(el);
      stack.push({ el, level: L, ordered: it.ordered });
    }
    while (top().level < L) {
      let parentLi = top().el.lastElementChild as HTMLElement | null;
      if (!parentLi || parentLi.tagName !== 'LI') { parentLi = doc.createElement('li'); top().el.appendChild(parentLi); }
      const el = make(it);
      parentLi.appendChild(el);
      stack.push({ el, level: top().level + 1, ordered: it.ordered });
    }
    top().el.appendChild(it.li);
  }
  return frag;
}

function flattenList(list: Element, depth: number, out: ListItemSpec[]) {
  const ordered = list.tagName === 'OL';
  const listType = olTypeFromStyle(parseStyle(list.getAttribute('style'))['list-style-type']) || list.getAttribute('type');
  const start = parseInt(list.getAttribute('start') || '', 10);
  let first = true;
  for (const child of Array.from(list.children)) {
    if (child.tagName === 'UL' || child.tagName === 'OL') { flattenList(child, depth + 1, out); continue; }
    if (child.tagName !== 'LI') continue;
    const li = child as HTMLElement;
    const nested = Array.from(li.children).filter(c => c.tagName === 'UL' || c.tagName === 'OL');
    nested.forEach(n => n.remove());
    const aria = parseInt(li.getAttribute('aria-level') || '', 10);
    const itemType = olTypeFromStyle(parseStyle(li.getAttribute('style'))['list-style-type']);
    out.push({
      level: aria > 0 ? aria : depth,
      ordered,
      olType: itemType || listType,
      li,
      start: first && isFinite(start) ? start : undefined,
    });
    first = false;
    nested.forEach(n => flattenList(n, depth + 1, out));
  }
}

const isWhitespaceNode = (n: Node | null) => !!n && n.nodeType === 3 && !(n.textContent || '').trim();

function nextElementSkippingSpace(n: Node): Element | null {
  let s = n.nextSibling;
  while (s && isWhitespaceNode(s)) s = s.nextSibling;
  return s && s.nodeType === 1 ? (s as Element) : null;
}

/** Rebuild Docs-style lists (flat items with aria-level, or ul directly inside ul) into nested lists. */
function rebuildHtmlLists(doc: Document, root: Element) {
  const isList = (e: Element | null) => !!e && (e.tagName === 'UL' || e.tagName === 'OL');
  const topLists = Array.from(root.querySelectorAll('ul,ol')).filter(l => !l.parentElement?.closest('ul,ol,li'));
  const done = new Set<Element>();
  for (const first of topLists) {
    if (done.has(first)) continue;
    const run: Element[] = [first];
    let nx = nextElementSkippingSpace(first);
    while (isList(nx)) { run.push(nx!); nx = nextElementSkippingSpace(nx!); }
    run.forEach(r => done.add(r));
    const items: ListItemSpec[] = [];
    run.forEach(l => flattenList(l, 1, items));
    if (!items.length) { run.forEach(r => r.remove()); continue; }
    const frag = buildLists(doc, items);
    first.parentNode!.insertBefore(frag, first);
    run.forEach(r => r.remove());
  }
}

function markerType(text: string): { ordered: boolean; olType: string | null } {
  const t = text.replace(/\u00a0/g, ' ').trim();
  if (/^\(?\d+[.)]?$/.test(t)) return { ordered: true, olType: '1' };
  if (/^\(?[ivxlcdm]{2,}[.)]$/.test(t) || /^\(?i[.)]$/.test(t)) return { ordered: true, olType: 'i' };
  if (/^\(?[IVXLCDM]{2,}[.)]$/.test(t) || /^\(?I[.)]$/.test(t)) return { ordered: true, olType: 'I' };
  if (/^\(?[a-z][.)]$/.test(t)) return { ordered: true, olType: 'a' };
  if (/^\(?[A-Z][.)]$/.test(t)) return { ordered: true, olType: 'A' };
  return { ordered: false, olType: null };
}

/** Word: <p style="mso-list:l0 level2 lfo1"> paragraphs -> nested ul/ol. */
function convertWordLists(doc: Document, root: Element) {
  const isListPara = (e: Element | null) => !!e && /^(P|H[1-6])$/.test(e.tagName) && /mso-list\s*:\s*l\d+/i.test(e.getAttribute('style') || '');
  const paras = Array.from(root.querySelectorAll('p,h1,h2,h3,h4,h5,h6')).filter(isListPara);
  const done = new Set<Element>();
  for (const first of paras) {
    if (done.has(first)) continue;
    const group: Element[] = [first];
    let nx = nextElementSkippingSpace(first);
    while (isListPara(nx)) { group.push(nx!); nx = nextElementSkippingSpace(nx!); }
    group.forEach(g => done.add(g));
    const items: ListItemSpec[] = group.map(p => {
      const st = parseStyle(p.getAttribute('style'));
      const level = parseInt(/level(\d+)/i.exec(st['mso-list'] || '')?.[1] || '1', 10);
      const markerEl = p.querySelector('[data-kb-listmarker]');
      const marker = markerEl?.textContent || '';
      markerEl?.remove();
      p.querySelectorAll('span').forEach(s => { if (/mso-list\s*:\s*ignore/i.test(s.getAttribute('style') || '')) s.remove(); });
      const { ordered, olType } = markerType(marker);
      const li = doc.createElement('li');
      const inner = doc.createElement('p');
      const keep: Record<string, string> = {};
      if (st['text-align']) keep['text-align'] = st['text-align'];
      if (st['line-height']) keep['line-height'] = st['line-height'];
      if (Object.keys(keep).length) inner.setAttribute('style', styleString(keep));
      while (p.firstChild) inner.appendChild(p.firstChild);
      li.appendChild(inner);
      return { level, ordered, olType, li };
    });
    first.parentNode!.insertBefore(buildLists(doc, items), first);
    group.forEach(g => g.remove());
  }
}

// ------------------------------------------------------------------ per-element conversion

interface Ctx { doc: Document; source: PasteSource; mode: 'paste' | 'load' }

function convertInline(el: HTMLElement, ctx: Ctx) {
  const st = parseStyle(el.getAttribute('style'));
  const inLink = !!el.closest('a');
  const tag = el.tagName;
  const wrappers: string[] = [];
  const weight = (st['font-weight'] || '').toLowerCase();
  if (weight === 'bold' || weight === 'bolder' || (/^\d+$/.test(weight) && parseInt(weight, 10) >= 600)) wrappers.push('strong');
  if ((st['font-style'] || '').toLowerCase() === 'italic') wrappers.push('em');
  const deco = `${st['text-decoration'] || ''} ${st['text-decoration-line'] || ''}`.toLowerCase();
  if (deco.includes('underline') && !inLink) wrappers.push('u');
  if (deco.includes('line-through')) wrappers.push('s');
  const va = (st['vertical-align'] || '').toLowerCase();
  const isScript = va === 'super' || va === 'sub';
  if (va === 'super') wrappers.push('sup');
  if (va === 'sub') wrappers.push('sub');

  const css: Record<string, string> = {};
  // colour
  let color = normalizeColor(st.color || (tag === 'FONT' ? el.getAttribute('color') : null));
  if (color && (BLACKS.has(color) || color === 'transparent' || (inLink && DEFAULT_LINK_COLORS.has(color)))) color = null;
  if (color) css.color = color;
  // highlight
  let bgRaw = st['background-color'] || st['mso-highlight'] || null;
  if (!bgRaw && st.background) bgRaw = st.background.split(/\s+/).find(p => !!normalizeColor(p)) || null;
  if (tag === 'MARK' && !bgRaw) bgRaw = '#fff3bf';
  const bg = normalizeColor(bgRaw);
  if (bg && !DROP_BG.has(bg)) css['background-color'] = bg;
  // font family -> catalog stack (source default fonts dropped so the theme governs)
  const fam = st['font-family'] || (tag === 'FONT' ? el.getAttribute('face') : null);
  if (fam) {
    const ff = firstFamily(fam);
    const isDefault = (ctx.source === 'docs' && DOCS_DEFAULT_FONTS.has(ff)) || (ctx.source === 'word' && WORD_DEFAULT_FONTS.has(ff));
    if (ff && !isDefault && !/^(inherit|initial|unset)$/.test(ff)) css['font-family'] = nearestFont(fam).stack;
  }
  // font size: pt -> px; 11pt/12pt document defaults dropped
  const fs = st['font-size'];
  if (fs && !isScript) {
    const m = /^(\d*\.?\d+)\s*(pt|px)$/i.exec(fs.trim());
    if (m) {
      const n = parseFloat(m[1]);
      const unit = m[2].toLowerCase();
      const isDefault = unit === 'pt' && (n === 11 || n === 12);
      if (!isDefault) {
        const px = Math.round(unit === 'pt' ? n * (4 / 3) : n);
        if (px >= 6 && px <= 96) css['font-size'] = `${px}px`;
      }
    }
  }

  // Semantic tags themselves
  if (tag === 'B' || tag === 'STRONG') { if (!(weight === 'normal' || weight === '400')) wrappers.unshift('strong'); }
  if (tag === 'I' || tag === 'EM') { if ((st['font-style'] || '') !== 'normal') wrappers.unshift('em'); }
  if (tag === 'U' || tag === 'INS') wrappers.unshift('u');
  if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') wrappers.unshift('s');
  if (tag === 'SUP') wrappers.unshift('sup');
  if (tag === 'SUB') wrappers.unshift('sub');

  const uniq = Array.from(new Set(wrappers));
  const doc = ctx.doc;
  const cls = tag === 'SPAN' ? kbClasses(el) : '';
  const outer: HTMLElement | null = Object.keys(css).length || cls ? doc.createElement('span') : null;
  if (outer && Object.keys(css).length) outer.setAttribute('style', styleString(css));
  if (outer && cls) outer.setAttribute('class', cls);
  let innermost: HTMLElement | null = outer;
  let root: HTMLElement | null = outer;
  for (const w of uniq) {
    const e = doc.createElement(w);
    if (innermost) innermost.appendChild(e); else root = e;
    innermost = e;
  }
  if (!root || !innermost) { unwrap(el); return; }
  while (el.firstChild) innermost.appendChild(el.firstChild);
  el.replaceWith(root);
}

function convertBlock(el: HTMLElement, ctx: Ctx): HTMLElement | null {
  let tag = el.tagName;
  if (/^H[4-6]$/.test(tag)) { el = rename(el, 'h3'); tag = 'H3'; }
  const st = parseStyle(el.getAttribute('style'));
  const out: Record<string, string> = {};
  const align = (st['text-align'] || el.getAttribute('align') || '').toLowerCase();
  if (/^(center|right|justify)$/.test(align)) out['text-align'] = align;

  if (tag === 'P' || /^H[1-3]$/.test(tag)) {
    const lh = normalizeLineHeight(st['line-height']);
    if (lh) out['line-height'] = lh;
    const m = marginParts(st);
    const top = lengthToPx(m.top);
    const bottom = lengthToPx(m.bottom);
    if (top && top > 0) out['margin-top'] = `${Math.min(MAX_SPACING_PX, Math.round(top))}px`;
    if (bottom && bottom > 0) out['margin-bottom'] = `${Math.min(MAX_SPACING_PX, Math.round(bottom))}px`;
    const left = lengthToPx(m.left);
    const step = ctx.mode === 'load' ? INDENT_PX : SOURCE_INDENT_PX;
    const existing = parseInt(el.getAttribute('data-indent') || '', 10);
    const indent = existing > 0 ? existing : left && left > 0 && !el.closest('li') ? Math.round(left / step) : 0;
    const cls = kbClasses(el);
    for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
    if (cls) el.setAttribute('class', cls);
    if (indent > 0) el.setAttribute('data-indent', String(Math.min(MAX_INDENT, indent)));
  } else if (tag === 'TD' || tag === 'TH') {
    const keep = ['colspan', 'rowspan', 'colwidth', 'class'].map(a => [a, el.getAttribute(a)] as const);
    let bgRaw = st['background-color'] || el.getAttribute('bgcolor');
    if (!bgRaw && st.background) bgRaw = st.background.split(/\s+/).find(p => !!normalizeColor(p)) || null;
    const bg = normalizeColor(bgRaw);
    if (bg && bg !== 'transparent' && bg !== '#ffffff' && bg !== 'white') out['background-color'] = bg;
    const widthPx = lengthToPx(st.width) ?? (el.getAttribute('width') && /^\d+$/.test(el.getAttribute('width')!) ? parseInt(el.getAttribute('width')!, 10) : null);
    for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
    keep.forEach(([k, v]) => { if (v) el.setAttribute(k, v); });
    if (!el.getAttribute('colwidth') && widthPx && widthPx > 0 && (parseInt(el.getAttribute('colspan') || '1', 10) || 1) === 1) {
      el.setAttribute('colwidth', String(Math.round(widthPx)));
    }
  } else {
    for (const a of Array.from(el.attributes)) if (a.name !== 'class') el.removeAttribute(a.name);
  }
  if (Object.keys(out).length) el.setAttribute('style', styleString(out));
  return el;
}

function convertLink(el: HTMLElement) {
  const raw = el.getAttribute('href');
  if (!raw) { unwrap(el); return; }
  const href = safeHref(unwrapGoogleRedirect(raw.trim()));
  if (!href) { unwrap(el); return; }
  const st = parseStyle(el.getAttribute('style'));
  const color = normalizeColor(st.color);
  const cls = el.getAttribute('class');
  for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
  el.setAttribute('href', href);
  if (cls) el.setAttribute('class', cls);
  if (color && !DEFAULT_LINK_COLORS.has(color) && !BLACKS.has(color) && color !== 'transparent') {
    // Link colour becomes a text colour inside the link (the serializer lifts it back onto <a>).
    const span = el.ownerDocument.createElement('span');
    span.setAttribute('style', `color:${color}`);
    while (el.firstChild) span.appendChild(el.firstChild);
    el.appendChild(span);
  }
}

function convertImage(el: HTMLElement, ctx: Ctx) {
  const src = (el.getAttribute('src') || '').trim();
  let rehost = false;
  if (/^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml);base64,/i.test(src)) rehost = true;
  else if (/^https?:\/\//i.test(src)) rehost = ctx.mode === 'paste' && !isHostedImage(src);
  else if (!src.startsWith('/api/uploads/')) { el.remove(); return; } // file:, cid:, blob: … cannot be fetched
  const st = parseStyle(el.getAttribute('style'));
  const wAttr = el.getAttribute('width');
  let width = wAttr && /^\d+(\.\d+)?$/.test(wAttr) ? parseFloat(wAttr) : lengthToPx(st.width);
  if (width && width > MAX_IMAGE_WIDTH) width = MAX_IMAGE_WIDTH;
  const alt = el.getAttribute('alt') || '';
  const align = el.getAttribute('data-align');
  for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
  el.setAttribute('src', src);
  el.setAttribute('alt', alt);
  if (width && width > 0) el.setAttribute('width', String(Math.round(width)));
  if (rehost) el.setAttribute('data-kb-rehost', '1');
  if (align) el.setAttribute('data-align', align);
}

// ------------------------------------------------------------------ fields

/** Wrap {Name} text in <span data-field="Name"> so the editor parses placeholder chips. */
export function fieldsToChips(root: Element) {
  const doc = root.ownerDocument;
  const walker = doc.createTreeWalker(root, 4 /* NodeFilter.SHOW_TEXT */);
  const texts: Text[] = [];
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if ((n.parentElement?.closest('[data-field],style,script'))) continue;
    if (fieldRegex('').test(n.textContent || '')) texts.push(n as Text);
  }
  for (const t of texts) {
    const s = t.textContent || '';
    const frag = doc.createDocumentFragment();
    let last = 0;
    for (const m of s.matchAll(fieldRegex('g'))) {
      if (m.index! > last) frag.appendChild(doc.createTextNode(s.slice(last, m.index)));
      const span = doc.createElement('span');
      span.setAttribute('data-field', m[1].trim());
      span.textContent = m[0];
      frag.appendChild(span);
      last = m.index! + m[0].length;
    }
    if (last < s.length) frag.appendChild(doc.createTextNode(s.slice(last)));
    t.replaceWith(frag);
  }
}

// ------------------------------------------------------------------ main entry

export function normalizePastedHtml(html: string, opts: NormalizeOptions = {}): string {
  const mode = opts.mode ?? 'paste';
  const source = detectSource(html);
  const doc = new DOMParser().parseFromString(preClean(html), 'text/html');
  const body = doc.body;
  const ctx: Ctx = { doc, source, mode };

  // 1. junk
  body.querySelectorAll('script,style,meta,link,title,noscript,iframe,object,embed,svg,template,xml,input,button,select,textarea,video,audio,canvas').forEach(e => e.remove());
  const cw = doc.createTreeWalker(body, 128 /* NodeFilter.SHOW_COMMENT */);
  const comments: Node[] = [];
  for (let n = cw.nextNode(); n; n = cw.nextNode()) comments.push(n);
  comments.forEach(c => c.parentNode?.removeChild(c));

  // 2. wrappers
  body.querySelectorAll('b[id^="docs-internal-guid"]').forEach(unwrap);
  body.querySelectorAll('b,strong').forEach(b => {
    const w = (parseStyle(b.getAttribute('style'))['font-weight'] || '').toLowerCase();
    if (w === 'normal' || w === '400') unwrap(b);
  });
  // Our own serialized aligned images: <table class="kb-image"><tr><td align=center><img>
  body.querySelectorAll('table.kb-image').forEach(t => {
    const img = t.querySelector('img');
    const td = t.querySelector('td');
    if (img) { img.setAttribute('data-align', td?.getAttribute('align') || 'left'); t.replaceWith(img); } else t.remove();
  });

  // 3. lists
  convertWordLists(doc, body);
  rebuildHtmlLists(doc, body);

  // 4. element pass (deepest first)
  const all = Array.from(body.querySelectorAll('*')).reverse() as HTMLElement[];
  for (const el of all) {
    if (!el.isConnected) continue;
    const tag = el.tagName;
    const kb = Array.from(el.classList).some(c => KB_CLASS.test(c));
    switch (tag) {
      case 'SPAN': case 'FONT': case 'MARK':
        if (el.hasAttribute('data-field')) break;
        convertInline(el, ctx); break;
      case 'B': case 'STRONG': case 'I': case 'EM': case 'U': case 'INS': case 'S': case 'STRIKE': case 'DEL': case 'SUP': case 'SUB':
        convertInline(el, ctx); break;
      case 'A': convertLink(el); break;
      case 'IMG': convertImage(el, ctx); break;
      case 'P': case 'H1': case 'H2': case 'H3': case 'H4': case 'H5': case 'H6': case 'TD': case 'TH':
        convertBlock(el, ctx); break;
      case 'LI': {
        // Serialized email lists carry the first paragraph's alignment/line-height on the <li>
        // with the text inline; restore it as a <p> so the editor keeps it.
        const st = parseStyle(el.getAttribute('style'));
        const keep: Record<string, string> = {};
        const al = (st['text-align'] || '').toLowerCase();
        if (/^(center|right|justify)$/.test(al)) keep['text-align'] = al;
        const lh = normalizeLineHeight(st['line-height']);
        if (lh) keep['line-height'] = lh;
        for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
        const first = el.firstChild;
        if (first && !(first.nodeType === 1 && BLOCK_TAGS.has((first as Element).tagName))) {
          const para = doc.createElement('p');
          while (el.firstChild && !(el.firstChild.nodeType === 1 && BLOCK_TAGS.has((el.firstChild as Element).tagName))) para.appendChild(el.firstChild);
          el.insertBefore(para, el.firstChild);
        }
        const fp = el.firstElementChild;
        if (fp && fp.tagName === 'P' && Object.keys(keep).length && !fp.getAttribute('style')) fp.setAttribute('style', styleString(keep));
        break;
      }
      case 'UL': case 'OL': case 'TR': case 'TBODY': case 'THEAD': case 'TFOOT': case 'COLGROUP': case 'BLOCKQUOTE': case 'HR': case 'BR': {
        const keep: Record<string, string | null> = {
          start: tag === 'OL' ? el.getAttribute('start') : null,
          type: tag === 'OL' ? el.getAttribute('type') : null,
        };
        for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
        Object.entries(keep).forEach(([k, v]) => { if (v) el.setAttribute(k, v); });
        break;
      }
      case 'COL': {
        const w = el.getAttribute('width') || (lengthToPx(parseStyle(el.getAttribute('style')).width) ?? '') + '';
        for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
        if (w && /^\d+(\.\d+)?$/.test(w)) el.setAttribute('width', String(Math.round(parseFloat(w))));
        break;
      }
      case 'TABLE': {
        if (kb) {
          const keep = ['class', 'data-variant', 'align'].map(a => [a, el.getAttribute(a)] as const);
          for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
          keep.forEach(([k, v]) => { if (v) el.setAttribute(k, v); });
        } else {
          for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
        }
        break;
      }
      case 'DIV': {
        if (kb) { for (const a of Array.from(el.attributes)) if (!['class', 'data-variant'].includes(a.name)) el.removeAttribute(a.name); break; }
        if (hasBlockChild(el)) unwrap(el);
        else convertBlock(rename(el, 'p'), ctx);
        break;
      }
      case 'PRE': {
        const p = rename(el, 'p');
        for (const a of Array.from(p.attributes)) p.removeAttribute(a.name);
        p.innerHTML = escapeHtml(p.textContent || '').replace(/\r?\n/g, '<br>');
        break;
      }
      case 'CODE': case 'KBD': case 'SAMP': case 'SMALL': case 'BIG': case 'ABBR': case 'CITE': case 'Q': case 'LABEL': case 'TIME':
        unwrap(el); break;
      default:
        if (CONTAINER_TAGS.has(tag)) { if (hasBlockChild(el)) unwrap(el); else convertBlock(rename(el, 'p'), ctx); }
        else if (!['TABLE', 'TBODY', 'TR', 'TD', 'TH'].includes(tag)) unwrap(el);
    }
  }

  // 5. blank-line cleanup for clipboard content: Docs uses top-level <br> and Word uses
  // empty paragraphs for spacing; the theme's paragraph margins already provide it.
  if (mode === 'paste' && source !== 'html') {
    Array.from(body.childNodes).forEach(n => { if (n.nodeName === 'BR') n.parentNode?.removeChild(n); });
    body.querySelectorAll('p').forEach(p => { if (isBlankText(p) && !p.closest('td,th,li')) p.remove(); });
  }

  // 6. a paragraph holding only an image becomes just the (block) image, so the editor
  // doesn't split it into an image plus an empty paragraph.
  body.querySelectorAll('p > img').forEach(img => {
    const para = img.parentElement!;
    if (para.closest('li,td,th')) return;
    const others = Array.from(para.childNodes).filter(n => n !== img && !(n.nodeType === 3 && !(n.textContent || '').replace(/[\s\u00a0]/g, '')) && n.nodeName !== 'BR');
    if (!others.length) para.replaceWith(img);
  });

  // 7. final attribute allow-list
  body.querySelectorAll('*').forEach(el => {
    for (const a of Array.from(el.attributes)) {
      if (!KEEP_ATTRS.has(a.name)) el.removeAttribute(a.name);
      else if (a.name === 'class') {
        const kept = a.value.split(/\s+/).filter(c => KB_CLASS.test(c));
        if (kept.length) el.setAttribute('class', kept.join(' ')); else el.removeAttribute('class');
      }
    }
  });

  // 8. placeholders -> chips
  fieldsToChips(body);
  return body.innerHTML.trim();
}

/** Prepare stored/legacy HTML for loading into the rich editor. */
export function prepareHtmlForEditor(html: string): string {
  if (!html || !html.trim()) return '';
  // Full documents (legacy templates): use only the body.
  const bodyMatch = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
  return normalizePastedHtml(bodyMatch ? bodyMatch[1] : html, { mode: 'load' });
}

/**
 * Plain-text clipboard -> HTML. Blank lines separate paragraphs and single newlines become
 * line breaks; text without any blank line is treated as one paragraph per line.
 */
export function plainTextToHtml(text: string): string {
  const t = text.replace(/\r\n?/g, '\n').replace(/\t/g, '    ');
  const withChips = (s: string) => escapeHtml(s).replace(fieldRegex('g'), (_m, name: string) => `<span data-field="${name.trim()}">{${name.trim()}}</span>`);
  const hasBlank = /\n[ \u00a0]*\n/.test(t);
  const blocks = hasBlank ? t.split(/\n[ \u00a0]*\n+/) : t.split('\n');
  return blocks
    .map(b => b.replace(/^\n+|\n+$/g, ''))
    .filter((b, i, arr) => b.trim() || (i > 0 && i < arr.length - 1 && !hasBlank))
    .map(b => `<p>${withChips(b).replace(/\n/g, '<br>') || '<br>'}</p>`)
    .join('');
}
