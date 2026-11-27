// Data tables: TipTap table nodes plus a per-cell background colour.
import { Table, TableCell, TableHeader, TableRow } from '@tiptap/extension-table';
import { normalizeColor, parseStyle } from '../style-utils';

const backgroundColor = {
  default: null as string | null,
  parseHTML: (el: HTMLElement) => {
    const st = parseStyle(el.getAttribute('style'));
    const c = normalizeColor(st['background-color'] || el.getAttribute('bgcolor'));
    return c && c !== 'transparent' ? c : null;
  },
  renderHTML: (attrs: Record<string, any>) => (attrs.backgroundColor ? { style: `background-color: ${attrs.backgroundColor}` } : {}),
};

export const KbTableCell = TableCell.extend({
  addAttributes() {
    return { ...this.parent?.(), backgroundColor };
  },
});

export const KbTableHeader = TableHeader.extend({
  addAttributes() {
    return { ...this.parent?.(), backgroundColor };
  },
});

export const tableExtensions = [
  Table.configure({ resizable: false, HTMLAttributes: { class: 'kb-data-table' } }),
  TableRow,
  KbTableHeader,
  KbTableCell,
];
