// @vitest-environment jsdom
// Schema-level round trips with the real TipTap extension set (headless, no React).
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { Editor, generateJSON, type JSONContent } from '@tiptap/core';
import { buildExtensions } from '../src/editor/extensions';
import { normalizePastedHtml, prepareHtmlForEditor } from '../src/editor/paste';
import { serializeEmail } from '../src/editor/serialize';
import { starterSignatureHtml } from '../shared/themes';

const exts = buildExtensions({ headless: true });
const toJSON = (html: string) => generateJSON(html, exts) as JSONContent;
/** HTML -> editor JSON -> email HTML */
const compile = (html: string) => serializeEmail(toJSON(prepareHtmlForEditor(html)));

const find = (node: JSONContent, type: string): JSONContent[] => {
  const out: JSONContent[] = [];
  const walk = (n: JSONContent) => { if (n.type === type) out.push(n); n.content?.forEach(walk); };
  walk(node);
  return out;
};

describe('placeholder chips', () => {
  let editor: Editor | null = null;
  afterEach(() => { editor?.destroy(); editor = null; });

  it('loading HTML with {Name} creates chips and serializes back to {Name}', () => {
    const json = toJSON(prepareHtmlForEditor('<p>Dear {Name}, your dept is {Department}.</p>'));
    expect(find(json, 'field').map(f => f.attrs?.name)).toEqual(['Name', 'Department']);
    expect(serializeEmail(json)).toBe('<p>Dear {Name}, your dept is {Department}.</p>');
  });

  it('typing {Name} becomes a chip via the input rule', () => {
    editor = new Editor({ element: document.createElement('div'), extensions: exts, content: '<p>Hi</p>' });
    editor.commands.focus('end');
    const type = (text: string) => {
      for (const ch of text) {
        const { from, to } = editor!.state.selection;
        const handled = editor!.view.someProp('handleTextInput', f => f(editor!.view, from, to, ch, () => editor!.state.tr.insertText(ch, from, to)));
        if (!handled) editor!.view.dispatch(editor!.state.tr.insertText(ch, from, to));
      }
    };
    type(' {Roll No}');
    const fields = find(editor.getJSON(), 'field');
    expect(fields).toHaveLength(1);
    expect(fields[0].attrs?.name).toBe('Roll No');
    expect(serializeEmail(editor.getJSON())).toBe('<p>Hi {Roll No}</p>');
  });

  it('insertField command and getText render {Name}', () => {
    editor = new Editor({ element: document.createElement('div'), extensions: exts, content: '<p>A</p>' });
    editor.chain().focus('end').insertField('Email').run();
    expect(editor.getText()).toBe('A{Email}');
    expect(serializeEmail(editor.getJSON())).toBe('<p>A{Email}</p>');
  });
});

describe('custom blocks round-trip through the editor schema', () => {
  const blocks = [
    '<table role="presentation" class="kb-button" data-variant="secondary" align="center" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-button-cell"><a class="kb-button-link" href="https://kiit.ac.in/r" target="_blank">Register</a></td></tr></table>',
    '<table role="presentation" class="kb-callout" data-variant="success" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-callout-cell"><p>Done <strong>well</strong></p></td></tr></table>',
    '<table role="presentation" class="kb-details" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-details-label">Date</td><td class="kb-details-value">12 Dec</td></tr><tr><td class="kb-details-label">Venue</td><td class="kb-details-value">{Venue}</td></tr></table>',
    '<table role="presentation" class="kb-divider" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-divider-cell">&nbsp;</td></tr></table>',
  ];

  it.each(blocks)('serializer output parses back to the same block: %#', html => {
    expect(compile(html)).toBe(html);
  });

  it('parses blocks into the right node types (not data tables)', () => {
    const json = toJSON(prepareHtmlForEditor(blocks.join('')));
    expect(json.content!.map(n => n.type)).toEqual(['kbButton', 'kbCallout', 'kbDetails', 'kbDivider']);
    expect(find(json, 'table')).toHaveLength(0);
    expect(json.content![0].attrs).toEqual({ label: 'Register', href: 'https://kiit.ac.in/r', variant: 'secondary', align: 'center' });
  });

  it('editor renderHTML (clipboard copy) re-parses to identical JSON', () => {
    const editor = new Editor({ element: document.createElement('div'), extensions: buildExtensions({ headless: true }), content: toJSON(prepareHtmlForEditor(blocks.join(''))) });
    // Button/callout/details use React node views in the app; headless rendering uses renderHTML.
    const json1 = editor.getJSON();
    const json2 = toJSON(normalizePastedHtml(editor.getHTML(), { mode: 'load' }));
    expect(serializeEmail(json2)).toBe(serializeEmail(json1));
    editor.destroy();
  });

  it('a full formatted document is stable across serialize -> load -> serialize', () => {
    const first = compile(normalizePastedHtml(readFileSync(path.join(__dirname, 'fixtures/paste/google-docs.html'), 'utf8')));
    const second = compile(first);
    expect(second).toBe(first);
  });
});

describe('end-to-end paste -> email HTML', () => {
  it('Google Docs fixture compiles to email-safe markup', () => {
    const html = serializeEmail(toJSON(normalizePastedHtml(readFileSync(path.join(__dirname, 'fixtures/paste/google-docs.html'), 'utf8'))));
    expect(html).toContain('<h1 style="line-height:138%;mso-line-height-rule:exactly;margin-top:27px;margin-bottom:8px">');
    expect(html).toContain("font-family:'Playfair Display', Georgia, 'Times New Roman', serif;font-size:27px");
    expect(html).toContain('<p style="line-height:138%;mso-line-height-rule:exactly">Dear {Name},</p>');
    expect(html).toContain('margin-left:32px'); // indent level 1
    expect(html).toContain('<a href="https://kiit.ac.in/iqac/review-2026?ref=mail&amp;x=1" target="_blank" rel="noopener">IQAC review page</a>');
    expect(html).toMatch(/<ul style="[^"]+"><li style="[^"]+">Opening remarks<ul style="margin:6px 0 0 0;/);
    expect(html).toContain('<ol type="a" style="margin:6px 0 0 0;');
    expect(html).toContain('<table border="1" cellpadding="6" cellspacing="0"');
    expect(html).toContain('background-color:#cfe2f3');
    expect(html).toContain('width="180"');
    // The googleusercontent image is still pending rehost here; it is emitted with its width
    expect(html).toContain('width="600" style="display:block;max-width:100%;height:auto;border:0;"');
    expect(html).not.toMatch(/data:|data-kb|data-field|class="(?!kb-)/);
  });

  it('Word fixture compiles without mso junk and with real lists', () => {
    const html = serializeEmail(toJSON(normalizePastedHtml(readFileSync(path.join(__dirname, 'fixtures/paste/word.html'), 'utf8'))));
    expect(html).not.toMatch(/mso-(?!line-height-rule)|Mso|o:p|file:/);
    expect(html).toMatch(/<ul style="[^"]+"><li style="[^"]+">Day one: curriculum mapping<ul/);
    expect(html).toMatch(/<ol style="[^"]+"><li style="[^"]+">Register online<\/li>/);
    expect(html).toContain('background-color:yellow');
  });
});

describe('signatures', () => {
  it('starterSignatureHtml keeps its kb-signature classes through load + serialize', () => {
    const sig = starterSignatureHtml({ name: 'Dr. Asha Rao', title: 'Coordinator', email: 'iqac@kiit.ac.in', phone: '+91 674 000 0000' } as any);
    const out = compile(sig);
    for (const cls of ['kb-signature-closing', 'kb-signature-name', 'kb-signature-title', 'kb-signature-org', 'kb-signature-contact']) {
      expect(out).toContain(`class="${cls}"`);
    }
    if (sig.includes('kb-signature-sep')) expect(out).toContain('<span class="kb-signature-sep">|</span>');
    expect(compile(out)).toBe(out);
  });
});
