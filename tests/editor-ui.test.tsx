// @vitest-environment jsdom
// Smoke test for <EmailEditor> / <HtmlCodeEditor>: renders with React node views, toolbar,
// debounced onChange, contentKey reload and theme restyle without resetting content.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type { Editor } from '@tiptap/react';
import { TooltipProvider } from '@/components/ui/tooltip';
import { EmailEditor, HtmlCodeEditor, type EmailEditorProps, type EmailEditorValue } from '../src/editor';

(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

beforeAll(() => {
  // jsdom lacks layout APIs that ProseMirror / Radix touch.
  const rect = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} });
  Range.prototype.getClientRects = () => ({ length: 0, item: () => null, [Symbol.iterator]: [][Symbol.iterator] }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect = rect as unknown as () => DOMRect;
  if (!document.elementFromPoint) document.elementFromPoint = () => null;
  if (!('ResizeObserver' in globalThis)) (globalThis as any).ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
});

let root: Root | null = null;
let host: HTMLDivElement | null = null;
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
  root = null; host = null;
  vi.useRealTimers();
});

/** Let TipTap's React node-view portals render. */
const flush = () => act(async () => { await new Promise(r => setTimeout(r, 0)); });

function mount(ui: React.ReactNode) {
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  act(() => root!.render(<TooltipProvider>{ui}</TooltipProvider>));
  return host;
}

const BLOCKS = '<p>Dear {Name},</p>'
  + '<table role="presentation" class="kb-button" data-variant="primary" align="left" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-button-cell"><a class="kb-button-link" href="https://kiit.ac.in" target="_blank">Register</a></td></tr></table>'
  + '<table role="presentation" class="kb-callout" data-variant="info" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-callout-cell"><p>Note</p></td></tr></table>'
  + '<table role="presentation" class="kb-details" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-details-label">Date</td><td class="kb-details-value">12 Dec</td></tr></table>'
  + '<p><img src="https://example.com/a.png" alt="" width="300"></p>';

describe('<EmailEditor>', () => {
  function setup(extra: Partial<EmailEditorProps> = {}) {
    const changes: EmailEditorValue[] = [];
    let editor: Editor | null = null;
    const props: EmailEditorProps = {
      doc: null, initialHtml: BLOCKS, theme: 'heritage', placeholders: ['Name', 'Email'],
      onChange: v => changes.push(v), onEditor: e => { editor = e; }, ...extra,
    };
    const el = mount(<EmailEditor {...props} />);
    return { el, changes, props, get editor() { return editor!; } };
  }

  it('renders the toolbar, theme CSS, chips and block node views', async () => {
    const { el } = setup();
    await flush();
    for (const label of ['Undo', 'Bold', 'Italic', 'Underline', 'Link', 'Image', 'Table', 'Insert block', 'Insert recipient field', 'Clear formatting', 'Font family', 'Font size', 'Text colour', 'Highlight']) {
      expect(el.querySelector(`[aria-label="${label}"]`), label).not.toBeNull();
    }
    expect(el.querySelector('style')!.textContent).toContain('.kb-body');
    const pm = el.querySelector('.ProseMirror')!;
    expect(pm.classList.contains('kb-body')).toBe(true);
    expect(pm.querySelector('.kb-field')!.textContent).toBe('{Name}');
    expect(pm.querySelector('table.kb-button a.kb-button-link')!.textContent).toBe('Register');
    expect(pm.querySelector('table.kb-callout td.kb-callout-cell p')!.textContent).toBe('Note');
    expect(pm.querySelector('table.kb-details td.kb-details-label')!.textContent).toBe('Date');
    expect(pm.querySelector('.kb-image-node img')!.getAttribute('src')).toBe('https://example.com/a.png');
    expect(pm.querySelector('.kb-image-alt-hint')).not.toBeNull(); // empty alt prompts for alt text
  });

  it('signature variant shows the compact toolbar inside .kb-body > .kb-signature', () => {
    const { el } = setup({ variant: 'signature', initialHtml: '<p class="kb-signature-name"><strong>Dr. A</strong></p>' });
    expect(el.querySelector('[aria-label="Insert block"]')).toBeNull();
    expect(el.querySelector('[aria-label="Table"]')).toBeNull();
    expect(el.querySelector('[aria-label="Bold"]')).not.toBeNull();
    expect(el.querySelector('.kb-body > div > .ProseMirror.kb-signature p.kb-signature-name')).not.toBeNull();
  });

  it('emits debounced {doc, html} and does not reset when doc/theme props change', () => {
    vi.useFakeTimers();
    const s = setup();
    act(() => { s.editor.chain().focus('end').insertContent('Hello').run(); });
    expect(s.changes).toHaveLength(0);
    act(() => { vi.advanceTimersByTime(300); });
    expect(s.changes).toHaveLength(1);
    const v = s.changes[0];
    expect((v.doc as any).type).toBe('doc');
    expect(v.html).toContain('<p>Dear {Name},</p>');
    expect(v.html).toContain('class="kb-button"');
    expect(v.html).toContain('Hello');
    expect(v.html).not.toMatch(/data-(?!variant)|ProseMirror|kb-field|react-renderer/);

    // Parent feeds the doc back and switches theme: content and history stay.
    act(() => root!.render(<TooltipProvider><EmailEditor {...s.props} doc={v.doc} theme="slate" /></TooltipProvider>));
    expect(s.editor.getText()).toContain('Hello');
    expect(s.editor.can().undo()).toBe(true);
  });

  it('reloads content only when contentKey changes', () => {
    const s = setup({ contentKey: 'a' });
    act(() => root!.render(<TooltipProvider><EmailEditor {...s.props} contentKey="a" initialHtml="<p>Other</p>" /></TooltipProvider>));
    expect(JSON.stringify(s.editor.getJSON())).toContain('"kbButton"');
    act(() => root!.render(<TooltipProvider><EmailEditor {...s.props} contentKey="b" initialHtml="<p>Other {Email}</p>" /></TooltipProvider>));
    expect(s.editor.getText()).toBe('Other {Email}');
    expect(s.editor.can().undo()).toBe(false);
  });

  it('flushes a pending change on unmount', () => {
    vi.useFakeTimers();
    const s = setup();
    act(() => { s.editor.chain().focus('end').insertContent('Last words').run(); });
    act(() => root!.unmount());
    root = null;
    expect(s.changes.at(-1)?.html).toContain('Last words');
  });
});

describe('<HtmlCodeEditor>', () => {
  it('renders with a size counter and toolbar', () => {
    const el = mount(<HtmlCodeEditor value="<p>Hello {Name}</p>" onChange={() => {}} placeholders={['Name']} />);
    expect(el.querySelector('[aria-label="Format HTML"]')).not.toBeNull();
    expect(el.querySelector('[aria-label="Insert recipient field"]')).not.toBeNull();
    expect(el.textContent).toContain('19 characters');
    expect(el.querySelector('.cm-editor')).not.toBeNull();
  });
});
