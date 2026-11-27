// Keeps theme classes on inline spans (e.g. <span class="kb-signature-sep">|</span>) so
// starter signatures round-trip through the editor unchanged.
import { Mark } from '@tiptap/core';

const kbOnly = (v: string | null) => (v || '').split(/\s+/).filter(c => /^kb-[a-z-]+$/.test(c) && c !== 'kb-field').join(' ');

export const KbSpan = Mark.create({
  name: 'kbSpan',
  excludes: '',
  addAttributes() {
    return {
      class: {
        default: null,
        parseHTML: el => kbOnly(el.getAttribute('class')) || null,
        renderHTML: attrs => (attrs.class ? { class: attrs.class } : {}),
      },
    };
  },
  parseHTML() {
    return [{
      tag: 'span[class]',
      consuming: false,
      getAttrs: el => ((el as HTMLElement).hasAttribute('data-field') || !kbOnly((el as HTMLElement).getAttribute('class')) ? false : null),
    }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', HTMLAttributes, 0];
  },
});
