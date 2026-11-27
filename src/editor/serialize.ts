// Email serializer: TipTap/ProseMirror JSON -> email-safe HTML body fragment.
//
// Pure (no DOM, no TipTap runtime) so it is trivially unit-testable and could run on the
// server too. Rules:
//  - All user formatting becomes INLINE styles (font stack, px size, colour, highlight,
//    alignment, line-height + mso-line-height-rule:exactly). Theme CSS (inlined later by
//    juice) styles everything else; inline styles win.
//  - Custom blocks emit exactly the markup in shared/blocks.ts.
//  - Never emits data: URIs, scripts, editor classes or editor-only attributes.
import { BLOCK_CLASSES as C } from '../../shared/blocks';
import { INDENT_PX, MAX_IMAGE_WIDTH, TABLE_BORDER } from './constants';
import {
  escapeAttr, escapeHtml, normalizeColor, normalizeFontSize, normalizeLineHeight, safeFontFamily,
  safeHref, safeImageSrc, styleString,
} from './style-utils';

export interface PMMark { type: string; attrs?: Record<string, any> }
export interface PMNode {
  type: string;
  attrs?: Record<string, any>;
  content?: PMNode[];
  marks?: PMMark[];
  text?: string;
}

const PRESENTATION = 'role="presentation"';
const TABLE_RESET = 'cellpadding="0" cellspacing="0" border="0"';

// ------------------------------------------------------------------ inline content

/** Fixed nesting order so adjacent text nodes share wrappers (links outermost). */
const MARK_ORDER = ['link', 'kbSpan', 'textStyle', 'highlight', 'bold', 'italic', 'underline', 'strike', 'superscript', 'subscript', 'code'];
const markRank = (m: PMMark) => { const i = MARK_ORDER.indexOf(m.type); return i < 0 ? 99 : i; };

function sameMark(a: PMMark, b: PMMark): boolean {
  if (a.type !== b.type) return false;
  return JSON.stringify(a.attrs || {}) === JSON.stringify(b.attrs || {});
}

function textStyleCss(attrs: Record<string, any> = {}): string {
  const color = normalizeColor(attrs.color);
  const bg = normalizeColor(attrs.backgroundColor);
  return styleString({
    'font-family': safeFontFamily(attrs.fontFamily),
    'font-size': normalizeFontSize(attrs.fontSize),
    color: color && color !== 'transparent' ? color : null,
    'background-color': bg && bg !== 'transparent' ? bg : null,
  });
}

/** Opening and closing tags for a mark. `ctx` is the first inline node it wraps. */
function markTags(mark: PMMark, ctx: PMNode): [string, string] | null {
  const a = mark.attrs || {};
  switch (mark.type) {
    case 'bold': return ['<strong>', '</strong>'];
    case 'italic': return ['<em>', '</em>'];
    case 'underline': return ['<u>', '</u>'];
    case 'strike': return ['<s>', '</s>'];
    case 'superscript': return ['<sup>', '</sup>'];
    case 'subscript': return ['<sub>', '</sub>'];
    case 'code': return ['<code style="font-family:\'Courier New\', Courier, monospace;">', '</code>'];
    case 'highlight': {
      const bg = normalizeColor(a.color) || '#fff3bf';
      return [`<span style="background-color:${bg}">`, '</span>'];
    }
    case 'kbSpan': {
      const cls = kbClassAttr(a.class);
      return cls ? [`<span${cls}>`, '</span>'] : null;
    }
    case 'textStyle': {
      const css = textStyleCss(a);
      return css ? [`<span style="${escapeAttr(css)}">`, '</span>'] : null;
    }
    case 'link': {
      const href = safeHref(a.href);
      if (!href) return null;
      // Mail clients recolour links unless the <a> itself carries the colour, so lift a
      // user-chosen text colour onto the link. Otherwise leave colour to the theme.
      const ts = ctx.marks?.find(m => m.type === 'textStyle');
      const color = normalizeColor(ts?.attrs?.color);
      const style = color && color !== 'transparent' ? ` style="color:${color}"` : '';
      return [`<a href="${escapeAttr(href)}" target="_blank" rel="noopener"${style}>`, '</a>'];
    }
    default: return null;
  }
}

function inlineLeaf(node: PMNode): string {
  switch (node.type) {
    case 'text': {
      // Preserve runs of spaces (HTML would collapse them).
      return escapeHtml(node.text || '').replace(/\u00a0/g, '&nbsp;').replace(/ {2}/g, ' &nbsp;').replace(/\n/g, '<br>');
    }
    case 'hardBreak': return '<br>';
    case 'field': return escapeHtml(`{${String(node.attrs?.name ?? '').trim()}}`);
    case 'image': return serializeImage(node); // inline images (rare) are allowed too
    default: return node.content ? serializeInline(node.content) : '';
  }
}

export function serializeInline(nodes: PMNode[] | undefined): string {
  if (!nodes?.length) return '';
  let out = '';
  const open: { mark: PMMark; close: string }[] = [];
  for (const node of nodes) {
    const marks = [...(node.marks || [])].sort((a, b) => markRank(a) - markRank(b));
    // Keep the longest common prefix of currently open marks.
    let keep = 0;
    while (keep < open.length && keep < marks.length && sameMark(open[keep].mark, marks[keep])) keep++;
    while (open.length > keep) out += open.pop()!.close;
    for (let i = keep; i < marks.length; i++) {
      const tags = markTags(marks[i], node);
      // A mark that renders nothing still occupies its slot so prefixes stay aligned.
      open.push({ mark: marks[i], close: tags ? tags[1] : '' });
      if (tags) out += tags[0];
    }
    out += inlineLeaf(node);
  }
  while (open.length) out += open.pop()!.close;
  return out;
}

// ------------------------------------------------------------------ blocks

function lineHeightCss(v: unknown): Record<string, string> {
  const lh = normalizeLineHeight(v as string);
  if (!lh) return {};
  return { 'line-height': `${Math.round(parseFloat(lh) * 100)}%`, 'mso-line-height-rule': 'exactly' };
}

/** Inline style for paragraph-like blocks. `extra` is merged first (so user values win). */
function blockCss(attrs: Record<string, any> = {}, extra: Record<string, string> = {}): string {
  const align = attrs.textAlign && attrs.textAlign !== 'left' && /^(center|right|justify)$/.test(attrs.textAlign) ? attrs.textAlign : null;
  const indent = Math.max(0, Math.min(8, Number(attrs.indent) || 0));
  const before = Number(attrs.spaceBefore);
  const after = Number(attrs.spaceAfter);
  return styleString({
    ...extra,
    'text-align': align,
    ...lineHeightCss(attrs.lineHeight),
    'margin-top': isFinite(before) && attrs.spaceBefore != null ? `${Math.round(before)}px` : extra['margin-top'],
    'margin-bottom': isFinite(after) && attrs.spaceAfter != null ? `${Math.round(after)}px` : extra['margin-bottom'],
    'margin-left': indent ? `${indent * INDENT_PX}px` : extra['margin-left'],
  });
}

const styleAttr = (css: string) => (css ? ` style="${escapeAttr(css)}"` : '');

/** ` class="kb-…"` keeping only kb-* class names (e.g. kb-signature-name), or ''. */
export function kbClassAttr(v: unknown): string {
  const kept = String(v ?? '').split(/\s+/).filter(c => /^kb-[a-z-]+$/.test(c));
  return kept.length ? ` class="${kept.join(' ')}"` : '';
}

function isEmptyParagraph(n: PMNode): boolean {
  return n.type === 'paragraph' && (!n.content || n.content.every(c => c.type === 'text' && !(c.text || '').trim()));
}

function paragraph(node: PMNode, extra?: Record<string, string>): string {
  const inner = serializeInline(node.content);
  return `<p${kbClassAttr(node.attrs?.kbClass)}${styleAttr(blockCss(node.attrs, extra))}>${inner || '&nbsp;'}</p>`;
}

function heading(node: PMNode): string {
  const level = Math.max(1, Math.min(3, Number(node.attrs?.level) || 1));
  return `<h${level}${kbClassAttr(node.attrs?.kbClass)}${styleAttr(blockCss(node.attrs))}>${serializeInline(node.content) || '&nbsp;'}</h${level}>`;
}

// Matches the theme's list spacing; inline so Outlook (which ignores much of <style>) agrees.
const LIST_CSS = 'margin:0 0 18px 0;padding:0 0 0 26px;';
const NESTED_LIST_CSS = 'margin:6px 0 0 0;padding:0 0 0 26px;';
const LI_CSS = { margin: '0 0 6px 0', padding: '0' };

function list(node: PMNode, nested: boolean): string {
  const ordered = node.type === 'orderedList';
  const tag = ordered ? 'ol' : 'ul';
  let attrs = '';
  if (ordered) {
    const start = Number(node.attrs?.start);
    if (isFinite(start) && start > 1) attrs += ` start="${Math.round(start)}"`;
    const type = node.attrs?.type;
    if (typeof type === 'string' && /^[1aAiI]$/.test(type) && type !== '1') attrs += ` type="${type}"`;
  }
  const items = (node.content || []).map(li => listItem(li)).join('');
  return `<${tag}${attrs} style="${nested ? NESTED_LIST_CSS : LIST_CSS}">${items}</${tag}>`;
}

function listItem(li: PMNode): string {
  const kids = li.content || [];
  let liStyle = styleString(LI_CSS);
  let body = '';
  kids.forEach((child, i) => {
    if (child.type === 'paragraph' && i === 0) {
      // Unwrap the first paragraph so theme <p> margins don't apply inside the bullet.
      liStyle = blockCss({ ...child.attrs, indent: 0 }, LI_CSS);
      body += serializeInline(child.content) || '&nbsp;';
    } else if (child.type === 'paragraph') {
      body += paragraph(child, { margin: '6px 0 0 0' });
    } else if (child.type === 'bulletList' || child.type === 'orderedList') {
      body += list(child, true);
    } else {
      body += serializeBlock(child);
    }
  });
  return `<li style="${escapeAttr(liStyle)}">${body}</li>`;
}

function serializeImage(node: PMNode): string {
  const src = safeImageSrc(node.attrs?.src);
  if (!src) return ''; // data:, blob:, file: and anything unhosted are dropped, never emitted
  const alt = escapeAttr(String(node.attrs?.alt ?? ''));
  const w = Number(node.attrs?.width);
  const width = isFinite(w) && w > 0 ? ` width="${Math.min(MAX_IMAGE_WIDTH, Math.round(w))}"` : '';
  const img = `<img src="${escapeAttr(src)}" alt="${alt}"${width} style="display:block;max-width:100%;height:auto;border:0;">`;
  const align = node.attrs?.align;
  if (align === 'center' || align === 'right') {
    return `<table ${PRESENTATION} class="kb-image" width="100%" ${TABLE_RESET}><tr><td align="${align}">${img}</td></tr></table>`;
  }
  return img;
}

function cellContent(cell: PMNode): string {
  const kids = cell.content || [];
  return kids.map((k, i) => {
    if (k.type === 'paragraph') {
      if (kids.length === 1 && isEmptyParagraph(k)) return '&nbsp;';
      return paragraph(k, { margin: i === kids.length - 1 ? '0' : '0 0 6px 0' });
    }
    return serializeBlock(k);
  }).join('');
}

function table(node: PMNode): string {
  const rows = (node.content || []).map(row => {
    const cells = (row.content || []).map(cell => {
      const tag = cell.type === 'tableHeader' ? 'th' : 'td';
      const a = cell.attrs || {};
      const bg = normalizeColor(a.backgroundColor);
      const align = /^(center|right|justify)$/.test(a.align || '') ? a.align : 'left';
      const css = styleString({
        border: `1px solid ${TABLE_BORDER}`,
        padding: '6px',
        'vertical-align': 'top',
        'text-align': align,
        'background-color': bg && bg !== 'transparent' ? bg : null,
      });
      let attrs = '';
      if (Number(a.colspan) > 1) attrs += ` colspan="${Number(a.colspan)}"`;
      if (Number(a.rowspan) > 1) attrs += ` rowspan="${Number(a.rowspan)}"`;
      const cw = Array.isArray(a.colwidth) && a.colwidth.length === 1 ? Number(a.colwidth[0]) : NaN;
      if (isFinite(cw) && cw > 0) attrs += ` width="${Math.round(cw)}"`;
      return `<${tag}${attrs} style="${escapeAttr(css)}">${cellContent(cell)}</${tag}>`;
    }).join('');
    return `<tr>${cells}</tr>`;
  }).join('');
  return `<table border="1" cellpadding="6" cellspacing="0" width="100%" style="border-collapse:collapse;width:100%;border-color:${TABLE_BORDER};margin:0 0 16px 0;">${rows}</table>`;
}

// ------------------------------------------------------------------ custom blocks (shared/blocks.ts)

export function buttonHtml(attrs: Record<string, any> = {}): string {
  const variant = attrs.variant === 'secondary' ? 'secondary' : 'primary';
  const align = attrs.align === 'center' ? 'center' : 'left';
  const href = safeHref(attrs.href) || '#';
  const label = escapeHtml(String(attrs.label || 'Open link'));
  return `<table ${PRESENTATION} class="${C.button}" data-variant="${variant}" align="${align}" ${TABLE_RESET}>`
    + `<tr><td class="${C.buttonCell}"><a class="${C.buttonLink}" href="${escapeAttr(href)}" target="_blank">${label}</a></td></tr></table>`;
}

function callout(node: PMNode): string {
  const v = node.attrs?.variant;
  const variant = v === 'warn' || v === 'success' ? v : 'info';
  const inner = serializeBlocks(node.content);
  return `<table ${PRESENTATION} class="${C.callout}" data-variant="${variant}" width="100%" ${TABLE_RESET}>`
    + `<tr><td class="${C.calloutCell}">${inner}</td></tr></table>`;
}

function details(node: PMNode): string {
  const rows = (node.content || []).filter(r => r.type === 'kbDetailsRow').map(row => {
    const label = row.content?.find(c => c.type === 'kbDetailsLabel');
    const value = row.content?.find(c => c.type === 'kbDetailsValue');
    return `<tr><td class="${C.detailsLabel}">${serializeInline(label?.content) || '&nbsp;'}</td>`
      + `<td class="${C.detailsValue}">${serializeInline(value?.content) || '&nbsp;'}</td></tr>`;
  }).join('');
  return `<table ${PRESENTATION} class="${C.details}" width="100%" ${TABLE_RESET}>${rows}</table>`;
}

export const DIVIDER_HTML = `<table ${PRESENTATION} class="${C.divider}" width="100%" ${TABLE_RESET}><tr><td class="${C.dividerCell}">&nbsp;</td></tr></table>`;

// ------------------------------------------------------------------ dispatcher

export function serializeBlock(node: PMNode): string {
  switch (node.type) {
    case 'paragraph': return paragraph(node);
    case 'heading': return heading(node);
    case 'bulletList':
    case 'orderedList': return list(node, false);
    case 'listItem': return listItem(node);
    case 'blockquote': return `<blockquote>${serializeBlocks(node.content)}</blockquote>`;
    case 'horizontalRule': return '<hr style="border:0;border-top:1px solid #d5d9df;margin:20px 0;height:0;">';
    case 'image': return serializeImage(node);
    case 'table': return table(node);
    case 'kbButton': return buttonHtml(node.attrs);
    case 'kbCallout': return callout(node);
    case 'kbDetails': return details(node);
    case 'kbDivider': return DIVIDER_HTML;
    case 'text':
    case 'hardBreak':
    case 'field': return serializeInline([node]);
    default:
      // Unknown wrapper: keep its children rather than dropping text.
      return node.content ? serializeBlocks(node.content) : '';
  }
}

export function serializeBlocks(nodes: PMNode[] | undefined): string {
  return (nodes || []).map(serializeBlock).join('');
}

/**
 * Compile an editor document (TipTap `editor.getJSON()`) into the email body fragment.
 * Leading/trailing empty paragraphs are trimmed (TipTap keeps a trailing one for editing).
 */
export function serializeEmail(doc: unknown): string {
  const root = doc as PMNode | null;
  if (!root || typeof root !== 'object') return '';
  const blocks = [...(root.type === 'doc' ? root.content || [] : [root])];
  while (blocks.length && isEmptyParagraph(blocks[blocks.length - 1])) blocks.pop();
  while (blocks.length && isEmptyParagraph(blocks[0])) blocks.shift();
  return serializeBlocks(blocks);
}
