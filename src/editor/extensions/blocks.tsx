// Custom email blocks (shared/blocks.ts): Button, Callout, Details panel, Divider.
// In the editor they render with the same classes/table structure as the sent mail, so the
// theme CSS styles them identically; the serializer emits the exact contract markup.
import { useState } from 'react';
import { Node, mergeAttributes, type Editor } from '@tiptap/core';
import { Selection, TextSelection } from '@tiptap/pm/state';
import type { Node as PMNode, ResolvedPos } from '@tiptap/pm/model';
import { NodeViewContent, NodeViewWrapper, ReactNodeViewRenderer, type ReactNodeViewProps } from '@tiptap/react';
import { Minus, Pencil, Plus, Trash2 } from 'lucide-react';
import { BLOCK_CLASSES as C } from '../../../shared/blocks';
import { CALLOUT_VARIANTS, type ButtonAlign, type ButtonVariant, type CalloutVariant } from '../constants';
import { ButtonDialog, type ButtonAttrs } from '../dialogs';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    kbBlocks: {
      insertKbButton: (attrs: Partial<ButtonAttrs>) => ReturnType;
      insertKbCallout: (variant?: CalloutVariant) => ReturnType;
      insertKbDetails: (rows?: [string, string][]) => ReturnType;
      insertKbDivider: () => ReturnType;
      addKbDetailsRow: () => ReturnType;
      removeKbDetailsRow: () => ReturnType;
    };
  }
}

const TABLE_RESET = { role: 'presentation', cellpadding: '0', cellspacing: '0', border: '0' };
const REACT_TABLE = { role: 'presentation', cellPadding: 0, cellSpacing: 0, border: 0 } as const;
const hasClass = (el: HTMLElement, c: string) => el.classList.contains(c);

const ctlBtn = 'inline-flex h-7 items-center gap-1 rounded px-2 text-xs font-medium text-white/90 hover:bg-white/15 focus-visible:outline focus-visible:outline-2 focus-visible:outline-[var(--gold)]';

// ------------------------------------------------------------------ Button

function ButtonView({ node, updateAttributes, deleteNode, selected, editor }: ReactNodeViewProps) {
  const a = node.attrs as ButtonAttrs;
  const [open, setOpen] = useState(false);
  const editable = editor.isEditable;
  return (
    <NodeViewWrapper className={`kb-block-node${selected ? ' is-selected' : ''}`} data-kb-node="button">
      <table {...REACT_TABLE} className={C.button} data-variant={a.variant} align={a.align} onDoubleClick={() => editable && setOpen(true)}>
        <tbody><tr><td className={C.buttonCell}>
          <a className={C.buttonLink} href={a.href} target="_blank" rel="noopener noreferrer" onClick={e => e.preventDefault()}>{a.label}</a>
        </td></tr></tbody>
      </table>
      {selected && editable && (
        <div className="kb-block-controls" contentEditable={false} role="toolbar" aria-label="Button options">
          <span className="truncate px-1 text-[11px] text-white/70" title={a.href}>{a.href}</span>
          <button type="button" className={ctlBtn} onClick={() => setOpen(true)} aria-label="Edit button"><Pencil className="size-3.5" aria-hidden /> Edit</button>
          <button type="button" className={ctlBtn} onClick={deleteNode} aria-label="Remove button"><Trash2 className="size-3.5" aria-hidden /></button>
        </div>
      )}
      <ButtonDialog open={open} onOpenChange={setOpen} isEditing initial={a} onSubmit={v => { updateAttributes(v); setOpen(false); }} />
    </NodeViewWrapper>
  );
}

export const KbButton = Node.create({
  name: 'kbButton',
  group: 'block',
  atom: true,
  selectable: true,
  draggable: true,

  addAttributes() {
    return {
      label: { default: 'Open link' },
      href: { default: 'https://' },
      variant: { default: 'primary' as ButtonVariant },
      align: { default: 'left' as ButtonAlign },
    };
  },

  parseHTML() {
    return [{
      tag: `table.${C.button}`,
      priority: 1000,
      getAttrs: el => {
        const t = el as HTMLElement;
        const link = t.querySelector('a');
        return {
          label: (link?.textContent || t.textContent || 'Open link').trim(),
          href: link?.getAttribute('href') || 'https://',
          variant: t.getAttribute('data-variant') === 'secondary' ? 'secondary' : 'primary',
          align: t.getAttribute('align') === 'center' ? 'center' : 'left',
        };
      },
    }];
  },

  renderHTML({ node }) {
    const a = node.attrs;
    return ['table', { ...TABLE_RESET, class: C.button, 'data-variant': a.variant, align: a.align },
      ['tbody', ['tr', ['td', { class: C.buttonCell }, ['a', { class: C.buttonLink, href: a.href, target: '_blank' }, a.label]]]]];
  },

  addNodeView() {
    return ReactNodeViewRenderer(ButtonView);
  },

  addCommands() {
    return {
      insertKbButton: attrs => ({ commands }) => commands.insertContent({ type: this.name, attrs }),
    };
  },
});

// ------------------------------------------------------------------ Callout

function CalloutView({ node, updateAttributes, deleteNode, editor, selected }: ReactNodeViewProps) {
  const variant = (node.attrs.variant || 'info') as CalloutVariant;
  return (
    <NodeViewWrapper className={`kb-block-node kb-callout-node${selected ? ' is-selected' : ''}`} data-kb-node="callout">
      <table {...REACT_TABLE} className={C.callout} data-variant={variant} width="100%">
        <tbody><tr>
          <NodeViewContent<'td'> as="td" className={C.calloutCell} />
        </tr></tbody>
      </table>
      {editor.isEditable && (
        <div className="kb-block-controls kb-block-controls--hover" contentEditable={false} role="toolbar" aria-label="Callout options">
          {CALLOUT_VARIANTS.map(v => (
            <button key={v.id} type="button" className={`${ctlBtn}${v.id === variant ? ' bg-white/20' : ''}`} aria-pressed={v.id === variant} onClick={() => updateAttributes({ variant: v.id })}>{v.label}</button>
          ))}
          <button type="button" className={ctlBtn} onClick={deleteNode} aria-label="Remove callout"><Trash2 className="size-3.5" aria-hidden /></button>
        </div>
      )}
    </NodeViewWrapper>
  );
}

export const KbCallout = Node.create({
  name: 'kbCallout',
  group: 'block',
  content: '(paragraph | heading | bulletList | orderedList)+',
  defining: true,

  addAttributes() {
    return {
      variant: {
        default: 'info' as CalloutVariant,
        parseHTML: el => { const v = el.getAttribute('data-variant'); return v === 'warn' || v === 'success' ? v : 'info'; },
        renderHTML: attrs => ({ 'data-variant': attrs.variant }),
      },
    };
  },

  parseHTML() {
    return [
      { tag: `table.${C.callout}`, priority: 1000, contentElement: `td.${C.calloutCell}` },
      { tag: `div.${C.callout}`, priority: 1000, contentElement: (el: HTMLElement) => el.querySelector(`.${C.calloutCell}`) || el },
    ];
  },

  renderHTML({ HTMLAttributes }) {
    return ['table', mergeAttributes(TABLE_RESET, { class: C.callout, width: '100%' }, HTMLAttributes), ['tbody', ['tr', ['td', { class: C.calloutCell }, 0]]]];
  },

  addNodeView() {
    return ReactNodeViewRenderer(CalloutView);
  },

  addCommands() {
    return {
      insertKbCallout: (variant = 'info') => ({ commands, state }) => {
        const { $from, empty } = state.selection;
        // Wrap the current paragraph(s) when they have text; otherwise insert a fresh callout.
        if (!empty || $from.parent.textContent.trim()) {
          if (commands.wrapIn(this.name, { variant })) return true;
        }
        return commands.insertContent({ type: this.name, attrs: { variant }, content: [{ type: 'paragraph' }] });
      },
    };
  },
});

// ------------------------------------------------------------------ Details panel

function detailsContext($pos: ResolvedPos) {
  for (let d = $pos.depth; d > 0; d--) {
    const n = $pos.node(d);
    if (n.type.name === 'kbDetailsRow') return { rowDepth: d, row: n, details: $pos.node(d - 1), detailsDepth: d - 1 };
  }
  return null;
}

function DetailsView({ node, deleteNode, editor, getPos, selected }: ReactNodeViewProps) {
  const addRow = () => {
    const pos = typeof getPos === 'function' ? getPos() : undefined;
    if (typeof pos !== 'number') return;
    const end = pos + node.nodeSize - 1;
    const row = editor.schema.nodes.kbDetailsRow.createAndFill()!;
    editor.chain().command(({ tr }) => {
      tr.insert(end, row);
      tr.setSelection(TextSelection.create(tr.doc, end + 2));
      return true;
    }).focus().run();
  };
  const removeRow = () => {
    if (!editor.commands.removeKbDetailsRow()) {
      const pos = typeof getPos === 'function' ? getPos() : undefined;
      if (typeof pos !== 'number' || node.childCount <= 1) return;
      const last = node.lastChild!;
      const end = pos + node.nodeSize - 1;
      editor.chain().command(({ tr }) => { tr.delete(end - last.nodeSize, end); return true; }).focus().run();
    }
  };
  return (
    <NodeViewWrapper className={`kb-block-node kb-details-node${selected ? ' is-selected' : ''}`} data-kb-node="details">
      <NodeViewContent<'table'> as="table" className={C.details} {...REACT_TABLE} width="100%" />
      {editor.isEditable && (
        <div className="kb-block-controls kb-block-controls--hover" contentEditable={false} role="toolbar" aria-label="Details panel options">
          <button type="button" className={ctlBtn} onClick={addRow}><Plus className="size-3.5" aria-hidden /> Row</button>
          <button type="button" className={ctlBtn} onClick={removeRow} disabled={node.childCount <= 1} aria-label="Remove row"><Minus className="size-3.5" aria-hidden /> Row</button>
          <button type="button" className={ctlBtn} onClick={deleteNode} aria-label="Remove details panel"><Trash2 className="size-3.5" aria-hidden /></button>
        </div>
      )}
    </NodeViewWrapper>
  );
}

export const KbDetails = Node.create({
  name: 'kbDetails',
  group: 'block',
  content: 'kbDetailsRow+',
  defining: true,
  isolating: true,

  parseHTML() {
    return [
      { tag: `table.${C.details}`, priority: 1000 },
      { tag: `div.${C.details}`, priority: 1000 },
    ];
  },

  renderHTML() {
    return ['table', { ...TABLE_RESET, class: C.details, width: '100%' }, ['tbody', 0]];
  },

  addNodeView() {
    return ReactNodeViewRenderer(DetailsView, { contentDOMElementTag: 'tbody' });
  },

  addCommands() {
    return {
      insertKbDetails: (rows = [['Date', ''], ['Time', ''], ['Venue', '']]) => ({ commands }) => commands.insertContent({
        type: this.name,
        content: rows.map(([l, v]) => ({
          type: 'kbDetailsRow',
          content: [
            { type: 'kbDetailsLabel', content: l ? [{ type: 'text', text: l }] : [] },
            { type: 'kbDetailsValue', content: v ? [{ type: 'text', text: v }] : [] },
          ],
        })),
      }),
      addKbDetailsRow: () => ({ state, tr, dispatch }) => {
        const ctx = detailsContext(state.selection.$from);
        if (!ctx) return false;
        const after = state.selection.$from.after(ctx.rowDepth);
        if (dispatch) {
          tr.insert(after, state.schema.nodes.kbDetailsRow.createAndFill()!);
          tr.setSelection(TextSelection.create(tr.doc, after + 2));
          dispatch(tr.scrollIntoView());
        }
        return true;
      },
      removeKbDetailsRow: () => ({ state, tr, dispatch }) => {
        const ctx = detailsContext(state.selection.$from);
        if (!ctx || ctx.details.childCount <= 1) return false;
        const $f = state.selection.$from;
        const from = $f.before(ctx.rowDepth);
        const to = $f.after(ctx.rowDepth);
        if (dispatch) {
          tr.delete(from, to);
          tr.setSelection(Selection.near(tr.doc.resolve(Math.max(0, from - 2)), -1));
          dispatch(tr.scrollIntoView());
        }
        return true;
      },
    };
  },
});

const cellKeys = (editor: Editor) => {
  const go = (dir: 1 | -1) => {
    const { state } = editor;
    const $f = state.selection.$from;
    const ctx = detailsContext($f);
    if (!ctx) return false;
    const inLabel = $f.parent.type.name === 'kbDetailsLabel';
    if (dir === 1 && inLabel) return editor.commands.setTextSelection($f.after() + 1);
    if (dir === -1 && !inLabel) return editor.commands.setTextSelection($f.before() - 1);
    // Value -> next row's label (or label -> previous row's value).
    const target = dir === 1 ? $f.after(ctx.rowDepth) + 2 : $f.before(ctx.rowDepth) - 2;
    if (target <= 0 || target >= state.doc.content.size) return false;
    const $t = state.doc.resolve(target);
    if (!detailsContext($t)) return false;
    return editor.commands.setTextSelection(target);
  };
  return {
    Enter: () => {
      const $f = editor.state.selection.$from;
      if (!detailsContext($f)) return false;
      if ($f.parent.type.name === 'kbDetailsLabel') return editor.commands.setTextSelection($f.after() + 1);
      return editor.commands.addKbDetailsRow();
    },
    Tab: () => go(1),
    'Shift-Tab': () => go(-1),
    Backspace: () => {
      const { selection } = editor.state;
      const $f = selection.$from;
      const ctx = detailsContext($f);
      if (!ctx || !selection.empty || $f.parentOffset !== 0) return false;
      if ($f.parent.type.name === 'kbDetailsValue') return editor.commands.setTextSelection($f.before() - 1); // jump back to label end
      if (ctx.row.textContent.trim() === '') return editor.commands.removeKbDetailsRow();
      return true; // never merge label into the previous row
    },
  };
};

export const KbDetailsRow = Node.create({
  name: 'kbDetailsRow',
  content: 'kbDetailsLabel kbDetailsValue',
  priority: 200,
  parseHTML() {
    return [
      { tag: 'tr', priority: 1000, context: 'kbDetails/' },
      { tag: 'div.kb-details-row', priority: 1000 },
    ];
  },
  renderHTML() { return ['tr', 0]; },
  addKeyboardShortcuts() { return cellKeys(this.editor); },
});

const cellNode = (name: 'kbDetailsLabel' | 'kbDetailsValue', cls: string) => Node.create({
  name,
  content: 'inline*',
  isolating: true,
  parseHTML() {
    return [
      { tag: `td.${cls}`, priority: 1000 },
      { tag: `div.${cls}`, priority: 1000 },
      { tag: 'td', priority: 900, context: 'kbDetailsRow/', getAttrs: (el: HTMLElement) => (hasClass(el, cls) || !el.className ? null : false) },
    ];
  },
  renderHTML() { return ['td', { class: cls }, 0]; },
});

export const KbDetailsLabel = cellNode('kbDetailsLabel', C.detailsLabel);
export const KbDetailsValue = cellNode('kbDetailsValue', C.detailsValue);

// ------------------------------------------------------------------ Divider

export const KbDivider = Node.create({
  name: 'kbDivider',
  group: 'block',
  atom: true,
  selectable: true,
  draggable: true,
  parseHTML() {
    return [{ tag: `table.${C.divider}`, priority: 1000 }];
  },
  renderHTML() {
    return ['table', { ...TABLE_RESET, class: C.divider, width: '100%' }, ['tbody', ['tr', ['td', { class: C.dividerCell }, '\u00a0']]]];
  },
  addCommands() {
    return {
      insertKbDivider: () => ({ commands }) => commands.insertContent({ type: this.name }),
    };
  },
});

export const blockExtensions = [KbButton, KbCallout, KbDetails, KbDetailsRow, KbDetailsLabel, KbDetailsValue, KbDivider];

/** True when `node` is one of the custom kb blocks (used by toolbar state). */
export const isKbBlock = (node: PMNode) => node.type.name.startsWith('kb');
