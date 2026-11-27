// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { detectSource, normalizePastedHtml, plainTextToHtml, prepareHtmlForEditor, unwrapGoogleRedirect } from '../src/editor/paste';
import { FONT_CATALOG } from '../shared/fonts';

const fixture = (f: string) => readFileSync(path.join(__dirname, 'fixtures/paste', f), 'utf8');
const stack = (id: string) => FONT_CATALOG.find(f => f.id === id)!.stack;
const parse = (html: string) => {
  const d = document.createElement('div');
  d.innerHTML = html;
  return d;
};

describe('Google Docs paste', () => {
  const raw = fixture('google-docs.html');
  const out = normalizePastedHtml(raw);
  const dom = parse(out);

  it('detects the source and unwraps the docs-internal-guid wrapper', () => {
    expect(detectSource(raw)).toBe('docs');
    expect(out).not.toMatch(/docs-internal-guid/);
    expect(dom.querySelector('b')).toBeNull();
    expect(dom.firstElementChild?.tagName).toBe('H1');
  });

  it('maps fonts to catalog stacks and drops the Docs default (Arial)', () => {
    const h1Span = dom.querySelector('h1 span')!;
    expect(h1Span.getAttribute('style')).toContain(`font-family:${stack('playfair')}`);
    expect(out).toContain(`font-family:${stack('georgia')}`);
    expect(out).toContain(`font-family:${stack('lato')}`);
    // Comfortaa is not in the catalog: mapped by category to a universal stack
    expect(dom.querySelector('p[style*="center"] span')!.getAttribute('style')).toContain(`font-family:${stack('arial')}`);
    // Ordinary Arial body text carries no font-family at all (theme governs)
    expect(dom.querySelector('p[style*="justify"]')!.innerHTML).not.toMatch(/font-family/);
  });

  it('converts pt to px and drops 11pt/12pt defaults', () => {
    expect(out).not.toMatch(/:\s*[\d.]+pt/); // no pt values left in any style
    expect(out).toContain('font-size:27px'); // 20pt
    expect(out).toContain('font-size:19px'); // 14pt
    expect(out).toContain('font-size:13px'); // 10pt
    expect(out).toContain('font-size:21px'); // 16pt heading
    expect(out).not.toContain('font-size:15px'); // 11pt dropped
    expect(out).not.toContain('font-size:16px'); // 12pt dropped
  });

  it('maps span styles to semantic marks and keeps only real colours', () => {
    expect(dom.querySelector('strong')?.textContent).toBe('Annual Quality Review 2026');
    expect(out).toContain('<strong>annual review meeting</strong>');
    expect(out).toContain('<em><u>12 December 2026</u></em>');
    expect(out).toContain('<s>Campus 3</s>');
    expect(out).toContain('2<sup>nd</sup>');
    expect(out).toContain('H<sub>2</sub>O');
    expect(out).toContain('color:#990000;background-color:#fff2cc');
    expect(out).not.toMatch(/#000000|transparent|white-space|font-weight|vertical-align|font-variant/);
  });

  it('keeps paragraph alignment, line-height, spacing and indent', () => {
    expect(dom.querySelector('p[style*="text-align:justify"]')).not.toBeNull();
    expect(dom.querySelector('p[style*="text-align:center"]')).not.toBeNull();
    expect(out).toContain('line-height:1.38');
    const indented = dom.querySelector('p[data-indent]')!;
    expect(indented.getAttribute('data-indent')).toBe('1'); // 36pt
    expect(indented.getAttribute('style')).toContain('line-height:1.5');
    expect(indented.getAttribute('style')).toContain('margin-bottom:11px'); // 8pt
    expect(dom.querySelector('h1')!.getAttribute('style')).toContain('margin-top:27px');
  });

  it('rebuilds nested bullet and numbered lists, keeping list types', () => {
    const ul = dom.querySelector(':scope > ul')!;
    expect(ul.children).toHaveLength(2);
    const lvl2 = ul.querySelector(':scope > li > ul')!;
    expect(lvl2.children).toHaveLength(2);
    expect(lvl2.querySelector(':scope > li:nth-child(2) > ul > li')!.textContent).toBe('NIRF and NAAC metrics');
    expect(ul.querySelector(':scope > li:last-child')!.textContent).toBe('Department presentations');
    // No invalid ul-directly-inside-ul left over
    expect(dom.querySelector('ul > ul, ol > ol, ul > ol, ol > ul')).toBeNull();
    const ol = dom.querySelector(':scope > ol')!;
    expect(ol.children).toHaveLength(2);
    const sub = ol.querySelector(':scope > li > ol')!;
    expect(sub.getAttribute('type')).toBe('a');
    expect(sub.children).toHaveLength(2);
  });

  it('keeps tables with cell backgrounds and column widths', () => {
    const table = dom.querySelector('table')!;
    expect(table.attributes).toHaveLength(0);
    expect(Array.from(table.querySelectorAll('col')).map(c => c.getAttribute('width'))).toEqual(['180', '422']);
    const cells = table.querySelectorAll('td');
    expect(cells).toHaveLength(4);
    expect(cells[0].getAttribute('style')).toBe('background-color:#cfe2f3');
    expect(cells[2].getAttribute('style')).toBeNull();
    expect(table.parentElement).toBe(dom); // the Docs <div align> wrapper is gone
  });

  it('unwraps google.com/url redirects and drops the default link colour', () => {
    const a = dom.querySelector('a')!;
    expect(a.getAttribute('href')).toBe('https://kiit.ac.in/iqac/review-2026?ref=mail&x=1');
    expect(a.innerHTML).toBe('IQAC review page');
  });

  it('marks pasted googleusercontent images for rehosting and caps the width', () => {
    const img = dom.querySelector('img')!;
    expect(img.getAttribute('data-kb-rehost')).toBe('1');
    expect(img.getAttribute('width')).toBe('600');
    expect(img.getAttribute('alt')).toBe('');
    expect(img.hasAttribute('height')).toBe(false);
    expect(img.hasAttribute('style')).toBe(false);
  });

  it('turns {Name} into a placeholder chip and drops Docs blank-line <br>s and editor attributes', () => {
    expect(dom.querySelector('[data-field="Name"]')?.textContent).toBe('{Name}');
    expect(dom.querySelector(':scope > br')).toBeNull();
    expect(out).not.toMatch(/\s(dir|role|aria-level|id)=/);
  });
});

describe('Word paste', () => {
  const raw = fixture('word.html');
  const out = normalizePastedHtml(raw);
  const dom = parse(out);

  it('strips mso junk, Office tags, comments, styles and Mso classes', () => {
    expect(detectSource(raw)).toBe('word');
    expect(out).not.toMatch(/mso-|<o:|<\/o:|Mso|<!--|<style|<meta|<xml|supportLists|v:shape|file:/i);
  });

  it('keeps formatting with mapped fonts, px sizes and highlight; drops Calibri/black defaults', () => {
    const title = dom.querySelector('p[style*="center"]')!;
    expect(title.querySelector('strong span')!.getAttribute('style')).toBe(`color:#1f3864;font-family:${stack('georgia')};font-size:21px`);
    expect(out).toContain('<em>three-day</em>');
    expect(out).toContain('<u>5 January</u>');
    expect(out).toMatch(/background-color:yellow/);
    expect(out).toContain('font-size:12px'); // 9pt
    expect(out).not.toMatch(/Calibri/);
    expect(dom.querySelector('[data-field="Name"]')).not.toBeNull();
  });

  it('converts mso-list paragraphs into nested lists of the right type', () => {
    const ul = dom.querySelector(':scope > ul')!;
    expect(ul.children).toHaveLength(2);
    expect(ul.querySelector(':scope > li > ul > li')!.textContent).toBe('Programme outcomes');
    expect(ul.textContent).not.toMatch(/·/);
    const ol = dom.querySelector(':scope > ol')!;
    expect(Array.from(ol.children).map(li => li.textContent)).toEqual(['Register online', 'Bring your laptop']);
    expect(ol.textContent).not.toMatch(/1\.|2\./);
  });

  it('keeps table backgrounds and widths, indents, links; drops local file images and blank paragraphs', () => {
    const tds = dom.querySelectorAll('td');
    expect(tds[0].getAttribute('style')).toBe('background-color:#d9e2f3');
    expect(tds[0].getAttribute('colwidth')).toBe('301');
    expect(dom.querySelector('table')!.attributes).toHaveLength(0);
    expect(dom.querySelector('p[data-indent="1"]')!.textContent).toContain('Indented note');
    const a = dom.querySelector('a')!;
    expect(a.getAttribute('href')).toBe('https://kiit.ac.in/fdp');
    expect(a.innerHTML).toBe('registration link'); // #0563C1 default link colour dropped
    expect(dom.querySelector('img')).toBeNull();
    expect(Array.from(dom.querySelectorAll('p')).every(p => p.textContent!.replace(/\s/g, '').length > 0)).toBe(true);
  });
});

describe('generic safety and helpers', () => {
  it('removes scripts, event handlers, javascript: links and unsafe images', () => {
    const out = normalizePastedHtml('<p onclick="x()">Hi <a href="javascript:alert(1)">bad</a><script>alert(1)</script><img src="blob:abc"><img src="cid:img1"></p><style>p{}</style>');
    expect(out).toBe('<p>Hi bad</p>');
  });

  it('marks data: images for upload in both modes but keeps external images on load', () => {
    const data = 'data:image/png;base64,iVBORw0KGgo=';
    expect(normalizePastedHtml(`<p><img src="${data}"></p>`)).toContain('data-kb-rehost="1"');
    expect(prepareHtmlForEditor(`<p><img src="${data}"></p>`)).toContain('data-kb-rehost="1"');
    expect(prepareHtmlForEditor('<p><img src="https://example.com/a.png" width="300"></p>')).not.toContain('data-kb-rehost');
    expect(normalizePastedHtml('<p><img src="https://store.public.blob.vercel-storage.com/x.png"></p>')).not.toContain('data-kb-rehost');
  });

  it('preserves kb-* classes (signatures, blocks) and strips other classes', () => {
    const html = '<p class="kb-signature-name foo"><strong>Dr. A</strong></p><p class="kb-signature-contact">a@kiit.ac.in&nbsp;<span class="kb-signature-sep">|</span>&nbsp;123</p><p class="MsoNormal">x</p>';
    const out = prepareHtmlForEditor(html);
    expect(out).toContain('<p class="kb-signature-name"><strong>Dr. A</strong></p>');
    expect(out).toContain('<span class="kb-signature-sep">|</span>');
    expect(out).toContain('<p>x</p>');
  });

  it('extracts the body of full legacy documents', () => {
    expect(prepareHtmlForEditor('<!DOCTYPE html><html><head><style>p{}</style></head><body><div style="max-width:600px"><p>Hello</p></div></body></html>')).toBe('<p>Hello</p>');
  });

  it('unwraps google redirects only', () => {
    expect(unwrapGoogleRedirect('https://www.google.com/url?q=https://a.com/x&sa=D')).toBe('https://a.com/x');
    expect(unwrapGoogleRedirect('https://www.google.co.in/url?q=https://b.org&sa=D')).toBe('https://b.org');
    expect(unwrapGoogleRedirect('https://www.google.com/search?q=x')).toBe('https://www.google.com/search?q=x');
  });

  it('plain text keeps paragraphs, line breaks and fields', () => {
    expect(plainTextToHtml('Dear {Name},\n\nLine one\nline two\n\n\nBye')).toBe(
      '<p>Dear <span data-field="Name">{Name}</span>,</p><p>Line one<br>line two</p><p>Bye</p>',
    );
    expect(plainTextToHtml('a\nb')).toBe('<p>a</p><p>b</p>');
    expect(plainTextToHtml('<b>x</b> & y')).toBe('<p>&lt;b&gt;x&lt;/b&gt; &amp; y</p>');
  });
});
