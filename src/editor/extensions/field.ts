// Placeholder field: an inline atom shown as a chip, serialized as plain {Name} text.
import { InputRule, Node, PasteRule, mergeAttributes } from '@tiptap/core';
import { FIELD_SOURCE } from '../constants';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    field: {
      insertField: (name: string) => ReturnType;
    };
  }
}

export const Field = Node.create({
  name: 'field',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  draggable: false,

  addAttributes() {
    return {
      name: {
        default: '',
        parseHTML: el => (el.getAttribute('data-field') || '').trim(),
        renderHTML: attrs => ({ 'data-field': attrs.name }),
      },
    };
  },

  parseHTML() {
    return [{ tag: 'span[data-field]', priority: 100 }];
  },

  renderHTML({ node, HTMLAttributes }) {
    return ['span', mergeAttributes(HTMLAttributes, { class: 'kb-field', title: `Recipient field: ${node.attrs.name}` }), `{${node.attrs.name}}`];
  },

  renderText({ node }) {
    return `{${node.attrs.name}}`;
  },

  addCommands() {
    return {
      insertField: name => ({ commands }) => {
        const n = name.trim().replace(/[{}]/g, '');
        if (!n) return false;
        return commands.insertContent({ type: this.name, attrs: { name: n } });
      },
    };
  },

  addInputRules() {
    return [
      new InputRule({
        find: new RegExp(`${FIELD_SOURCE}$`),
        handler: ({ state, range, match }) => {
          const name = (match[1] || '').trim();
          if (!name) return null;
          const marks = state.selection.$from.marks();
          state.tr.replaceWith(range.from, range.to, this.type.create({ name }, null, marks));
        },
      }),
    ];
  },

  addPasteRules() {
    return [
      new PasteRule({
        find: new RegExp(FIELD_SOURCE, 'g'),
        handler: ({ state, range, match }) => {
          const name = (match[1] || '').trim();
          if (!name) return null;
          state.tr.replaceWith(range.from, range.to, this.type.create({ name }));
        },
      }),
    ];
  },
});
