// Paragraph/heading level formatting: line-height, indent levels, spacing before/after
// (kept from Docs/Word pastes) and kb-* classes (e.g. kb-signature-name).
import { Extension } from '@tiptap/core';
import type { EditorState, Transaction } from '@tiptap/pm/state';
import { INDENT_PX, MAX_INDENT } from '../constants';
import { lengthToPx, normalizeLineHeight, parseStyle } from '../style-utils';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    blockStyle: {
      /** Set (or clear with null) the line-height of the selected paragraphs/headings. */
      setBlockLineHeight: (lineHeight: string | null) => ReturnType;
      /** Indent: sinks list items, otherwise adds an indent level to paragraphs. */
      indentBlock: () => ReturnType;
      outdentBlock: () => ReturnType;
      /** Remove line-height, indent and spacing from the selected blocks. */
      clearBlockStyle: () => ReturnType;
    };
  }
}

const TYPES = ['paragraph', 'heading'];

function updateBlocks(state: EditorState, tr: Transaction, fn: (attrs: Record<string, any>) => Record<string, any> | null): boolean {
  const { from, to } = state.selection;
  let changed = false;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!TYPES.includes(node.type.name)) return true;
    const next = fn(node.attrs);
    if (next) { tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...next }); changed = true; }
    return false;
  });
  return changed;
}

export const BlockStyle = Extension.create({
  name: 'blockStyle',
  priority: 50, // after table/list keymaps so Tab in tables and lists keeps working

  addGlobalAttributes() {
    return [{
      types: TYPES,
      attributes: {
        lineHeight: {
          default: null,
          parseHTML: el => normalizeLineHeight(parseStyle(el.getAttribute('style'))['line-height']),
          renderHTML: attrs => (attrs.lineHeight ? { style: `line-height: ${attrs.lineHeight}` } : {}),
        },
        indent: {
          default: 0,
          parseHTML: el => {
            const d = parseInt(el.getAttribute('data-indent') || '', 10);
            if (d > 0) return Math.min(MAX_INDENT, d);
            const ml = lengthToPx(parseStyle(el.getAttribute('style'))['margin-left']);
            return ml && ml > 0 ? Math.min(MAX_INDENT, Math.round(ml / INDENT_PX)) : 0;
          },
          renderHTML: attrs => (attrs.indent ? { 'data-indent': attrs.indent, style: `margin-left: ${attrs.indent * INDENT_PX}px` } : {}),
        },
        spaceBefore: {
          default: null,
          parseHTML: el => {
            const v = lengthToPx(parseStyle(el.getAttribute('style'))['margin-top']);
            return v && v > 0 ? Math.round(v) : null;
          },
          renderHTML: attrs => (attrs.spaceBefore != null ? { style: `margin-top: ${attrs.spaceBefore}px` } : {}),
        },
        spaceAfter: {
          default: null,
          parseHTML: el => {
            const v = lengthToPx(parseStyle(el.getAttribute('style'))['margin-bottom']);
            return v && v > 0 ? Math.round(v) : null;
          },
          renderHTML: attrs => (attrs.spaceAfter != null ? { style: `margin-bottom: ${attrs.spaceAfter}px` } : {}),
        },
        kbClass: {
          default: null,
          parseHTML: el => (el.getAttribute('class') || '').split(/\s+/).filter(c => /^kb-[a-z-]+$/.test(c)).join(' ') || null,
          renderHTML: attrs => (attrs.kbClass ? { class: attrs.kbClass } : {}),
        },
      },
    }];
  },

  addCommands() {
    return {
      setBlockLineHeight: lineHeight => ({ state, tr, dispatch }) => {
        const lh = lineHeight ? normalizeLineHeight(lineHeight) : null;
        const ok = updateBlocks(state, tr, () => ({ lineHeight: lh }));
        if (ok && dispatch) dispatch(tr);
        return ok;
      },
      indentBlock: () => ({ state, tr, dispatch, editor, commands }) => {
        if (editor.isActive('listItem')) return commands.sinkListItem('listItem');
        const ok = updateBlocks(state, tr, a => ((a.indent || 0) < MAX_INDENT ? { indent: (a.indent || 0) + 1 } : null));
        if (ok && dispatch) dispatch(tr);
        return ok;
      },
      outdentBlock: () => ({ state, tr, dispatch, editor, commands }) => {
        if (editor.isActive('listItem')) return commands.liftListItem('listItem');
        const ok = updateBlocks(state, tr, a => ((a.indent || 0) > 0 ? { indent: a.indent - 1 } : null));
        if (ok && dispatch) dispatch(tr);
        return ok;
      },
      clearBlockStyle: () => ({ state, tr, dispatch }) => {
        const ok = updateBlocks(state, tr, () => ({ lineHeight: null, indent: 0, spaceBefore: null, spaceAfter: null }));
        if (ok && dispatch) dispatch(tr);
        return ok;
      },
    };
  },

  addKeyboardShortcuts() {
    return {
      Tab: () => this.editor.commands.indentBlock(),
      'Shift-Tab': () => this.editor.commands.outdentBlock(),
    };
  },
});
