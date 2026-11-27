import { describe, expect, it } from 'vitest';
import { serializeEmail, buttonHtml, DIVIDER_HTML, type PMNode } from '../src/editor/serialize';
import { FONT_CATALOG } from '../shared/fonts';

const doc = (...content: PMNode[]): PMNode => ({ type: 'doc', content });
const p = (content: PMNode[] = [], attrs: Record<string, any> = {}): PMNode => ({ type: 'paragraph', attrs, content });
const t = (text: string, marks: PMNode['marks'] = []): PMNode => ({ type: 'text', text, marks });

describe('serializeEmail: text and paragraphs', () => {
  it('emits inline styles for user formatting', () => {
    const lato = FONT_CATALOG.find(f => f.id === 'lato')!.stack;
    const html = serializeEmail(doc(p([
      t('Hello ', []),
      t('world', [{ type: 'bold' }, { type: 'textStyle', attrs: { fontFamily: lato, fontSize: '18px', color: '#1F3864', backgroundColor: '#fff3bf' } }]),
    ], { textAlign: 'center', lineHeight: '1.5' })));
    expect(html).toBe(
      `<p style="text-align:center;line-height:150%;mso-line-height-rule:exactly">Hello `
      + `<span style="font-family:Lato, 'Helvetica Neue', Arial, sans-serif;font-size:18px;color:#1f3864;background-color:#fff3bf"><strong>world</strong></span></p>`,
    );
  });

  it('shares wrappers between adjacent marked text and orders marks consistently', () => {
    const html = serializeEmail(doc(p([t('a', [{ type: 'italic' }, { type: 'bold' }]), t('b', [{ type: 'bold' }])])));
    expect(html).toBe('<p><strong><em>a</em>b</strong></p>');
  });

  it('renders sup, sub, underline, strike and line breaks', () => {
    const html = serializeEmail(doc(p([
      t('x', [{ type: 'superscript' }]), t('y', [{ type: 'subscript' }]), t('u', [{ type: 'underline' }]), t('s', [{ type: 'strike' }]),
      { type: 'hardBreak' }, t('end'),
    ])));
    expect(html).toBe('<p><sup>x</sup><sub>y</sub><u>u</u><s>s</s><br>end</p>');
  });

  it('writes indent and spacing as margins, headings h1-h3, and escapes text', () => {
    const html = serializeEmail(doc(
      { type: 'heading', attrs: { level: 2 }, content: [t('A & B <c>')] },
      p([t('In')], { indent: 2, spaceBefore: 12, spaceAfter: 4 }),
    ));
    expect(html).toBe('<h2>A &amp; B &lt;c&gt;</h2><p style="margin-top:12px;margin-bottom:4px;margin-left:64px">In</p>');
  });

  it('keeps kb-* classes on paragraphs (signature markup) and kb spans', () => {
    const html = serializeEmail(doc(
      p([t('Dr. A', [{ type: 'bold' }])], { kbClass: 'kb-signature-name' }),
      p([t('a@kiit.ac.in '), t('|', [{ type: 'kbSpan', attrs: { class: 'kb-signature-sep' } }]), t(' 123')], { kbClass: 'kb-signature-contact evil' }),
    ));
    expect(html).toBe('<p class="kb-signature-name"><strong>Dr. A</strong></p><p class="kb-signature-contact">a@kiit.ac.in <span class="kb-signature-sep">|</span> 123</p>');
  });

  it('trims leading/trailing empty paragraphs but keeps inner blank lines', () => {
    expect(serializeEmail(doc(p(), p([t('x')]), p(), p([t('y')]), p(), p()))).toBe('<p>x</p><p>&nbsp;</p><p>y</p>');
    expect(serializeEmail(doc(p()))).toBe('');
    expect(serializeEmail(null)).toBe('');
  });

  it('serializes placeholder chips as plain {Name}', () => {
    expect(serializeEmail(doc(p([t('Dear '), { type: 'field', attrs: { name: 'Name' } }, t(',')])))).toBe('<p>Dear {Name},</p>');
    expect(serializeEmail(doc(p([{ type: 'field', attrs: { name: 'Roll No' }, marks: [{ type: 'bold' }] }])))).toBe('<p><strong>{Roll No}</strong></p>');
  });
});

describe('serializeEmail: links', () => {
  it('adds target=_blank and only colours links the user coloured', () => {
    const link = { type: 'link', attrs: { href: 'https://kiit.ac.in' } };
    expect(serializeEmail(doc(p([t('site', [link])])))).toBe('<p><a href="https://kiit.ac.in" target="_blank" rel="noopener">site</a></p>');
    const html = serializeEmail(doc(p([t('red', [link, { type: 'textStyle', attrs: { color: '#7a1f2b' } }])])));
    expect(html).toBe('<p><a href="https://kiit.ac.in" target="_blank" rel="noopener" style="color:#7a1f2b"><span style="color:#7a1f2b">red</span></a></p>');
  });

  it('keeps one <a> across differently formatted text and drops unsafe hrefs', () => {
    const link = { type: 'link', attrs: { href: 'mailto:iqac@kiit.ac.in' } };
    expect(serializeEmail(doc(p([t('mail ', [link]), t('us', [link, { type: 'bold' }])])))).toBe('<p><a href="mailto:iqac@kiit.ac.in" target="_blank" rel="noopener">mail <strong>us</strong></a></p>');
    expect(serializeEmail(doc(p([t('x', [{ type: 'link', attrs: { href: 'javascript:alert(1)' } }])])))).toBe('<p>x</p>');
  });
});

describe('serializeEmail: lists, tables, images', () => {
  it('emits lists with inline margins and unwraps the first paragraph of each item', () => {
    const li = (text: string, nested?: PMNode): PMNode => ({ type: 'listItem', content: [p([t(text)]), ...(nested ? [nested] : [])] });
    const html = serializeEmail(doc({
      type: 'bulletList',
      content: [li('One', { type: 'orderedList', attrs: { start: 1, type: 'a' }, content: [li('Sub')] }), li('Two')],
    }));
    expect(html).toBe(
      '<ul style="margin:0 0 18px 0;padding:0 0 0 26px;">'
      + '<li style="margin:0 0 6px 0;padding:0">One<ol type="a" style="margin:6px 0 0 0;padding:0 0 0 26px;"><li style="margin:0 0 6px 0;padding:0">Sub</li></ol></li>'
      + '<li style="margin:0 0 6px 0;padding:0">Two</li></ul>',
    );
  });

  it('emits data tables with border/cellpadding/cellspacing, collapse and cell backgrounds', () => {
    const cell = (type: string, text: string, attrs: Record<string, any> = {}): PMNode => ({ type, attrs: { colspan: 1, rowspan: 1, colwidth: null, ...attrs }, content: [p([t(text)])] });
    const html = serializeEmail(doc({
      type: 'table',
      content: [
        { type: 'tableRow', content: [cell('tableHeader', 'Day', { backgroundColor: '#e6eef8', colwidth: [180] }), cell('tableHeader', 'Venue')] },
        { type: 'tableRow', content: [cell('tableCell', 'Mon'), cell('tableCell', 'Campus 6')] },
      ],
    }));
    expect(html).toContain('<table border="1" cellpadding="6" cellspacing="0" width="100%" style="border-collapse:collapse;');
    expect(html).not.toContain('role="presentation"');
    expect(html).toContain('<th width="180" style="border:1px solid #c9ced6;padding:6px;vertical-align:top;text-align:left;background-color:#e6eef8"><p style="margin:0">Day</p></th>');
    expect(html).toContain('<td style="border:1px solid #c9ced6;padding:6px;vertical-align:top;text-align:left"><p style="margin:0">Campus 6</p></td>');
  });

  it('emits the exact image markup, caps width, wraps aligned images, and never emits data: URIs', () => {
    expect(serializeEmail(doc({ type: 'image', attrs: { src: 'https://x.public.blob.vercel-storage.com/a.png', alt: 'Poster "2026"', width: 900 } })))
      .toBe('<img src="https://x.public.blob.vercel-storage.com/a.png" alt="Poster &quot;2026&quot;" width="600" style="display:block;max-width:100%;height:auto;border:0;">');
    expect(serializeEmail(doc({ type: 'image', attrs: { src: 'https://a.com/b.png', alt: '', width: 300, align: 'center' } })))
      .toBe('<table role="presentation" class="kb-image" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td align="center"><img src="https://a.com/b.png" alt="" width="300" style="display:block;max-width:100%;height:auto;border:0;"></td></tr></table>');
    const withData = serializeEmail(doc(
      { type: 'image', attrs: { src: 'data:image/png;base64,iVBORw0KGgo=', alt: 'x', rehost: true } },
      { type: 'image', attrs: { src: 'blob:http://localhost/123' } },
      p([t('after')]),
    ));
    expect(withData).toBe('<p>after</p>');
    expect(withData).not.toMatch(/data:/);
  });
});

describe('serializeEmail: custom blocks match shared/blocks.ts', () => {
  it('button', () => {
    expect(buttonHtml({ label: 'Register now', href: 'https://kiit.ac.in/r', variant: 'primary', align: 'center' })).toBe(
      '<table role="presentation" class="kb-button" data-variant="primary" align="center" cellpadding="0" cellspacing="0" border="0">'
      + '<tr><td class="kb-button-cell"><a class="kb-button-link" href="https://kiit.ac.in/r" target="_blank">Register now</a></td></tr></table>',
    );
    const html = serializeEmail(doc({ type: 'kbButton', attrs: { label: 'A <b>', href: 'javascript:x', variant: 'secondary', align: 'right' } }));
    expect(html).toContain('data-variant="secondary" align="left"');
    expect(html).toContain('href="#"');
    expect(html).toContain('>A &lt;b&gt;</a>');
  });

  it('callout', () => {
    const html = serializeEmail(doc({ type: 'kbCallout', attrs: { variant: 'warn' }, content: [p([t('Bring ID')]), p([t('Arrive early', [{ type: 'bold' }])])] }));
    expect(html).toBe(
      '<table role="presentation" class="kb-callout" data-variant="warn" width="100%" cellpadding="0" cellspacing="0" border="0">'
      + '<tr><td class="kb-callout-cell"><p>Bring ID</p><p><strong>Arrive early</strong></p></td></tr></table>',
    );
  });

  it('details', () => {
    const row = (l: string, v: PMNode[]): PMNode => ({ type: 'kbDetailsRow', content: [{ type: 'kbDetailsLabel', content: [t(l)] }, { type: 'kbDetailsValue', content: v }] });
    const html = serializeEmail(doc({ type: 'kbDetails', content: [row('Date', [t('12 Dec 2026')]), row('Venue', [{ type: 'field', attrs: { name: 'Venue' } }]), row('Time', [])] }));
    expect(html).toBe(
      '<table role="presentation" class="kb-details" width="100%" cellpadding="0" cellspacing="0" border="0">'
      + '<tr><td class="kb-details-label">Date</td><td class="kb-details-value">12 Dec 2026</td></tr>'
      + '<tr><td class="kb-details-label">Venue</td><td class="kb-details-value">{Venue}</td></tr>'
      + '<tr><td class="kb-details-label">Time</td><td class="kb-details-value">&nbsp;</td></tr></table>',
    );
  });

  it('divider', () => {
    expect(serializeEmail(doc({ type: 'kbDivider' }))).toBe(DIVIDER_HTML);
    expect(DIVIDER_HTML).toBe('<table role="presentation" class="kb-divider" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-divider-cell">&nbsp;</td></tr></table>');
  });

  it('never emits editor-only classes or attributes', () => {
    const html = serializeEmail(doc(
      p([t('x')], { kbClass: 'is-editor-empty ProseMirror' }),
      { type: 'image', attrs: { src: 'https://a.com/i.png', alt: '', rehost: true, width: 100 } },
    ));
    expect(html).not.toMatch(/ProseMirror|is-editor|data-kb-rehost|data-field|contenteditable|kb-field/);
  });
});
