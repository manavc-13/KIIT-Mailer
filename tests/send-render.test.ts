import { describe, it, expect } from 'vitest';
import { renderEmail, htmlToText } from '../server/lib/render/index.js';
import { hostedImageIds, inlineHostedImages, cidFor } from '../server/lib/send/images.js';
import type { ComposeContent } from '../shared/types.js';

const rich = (over: Partial<ComposeContent> = {}): ComposeContent => ({
  subject: 'Review meeting for {Name}', preheader: 'Agenda for {Name} inside', editorMode: 'rich',
  bodyHtml: '<p>Dear {Name},</p><p>The IQAC review meeting for your department is on Friday at 10:00 in the Campus 6 board room. Please bring the annual report.</p>',
  bodyText: null, theme: 'heritage', includeSignature: true, ...over,
});
const sender = { signatureHtml: '<p class="sig">Regards,<br>IQAC Office for {Name}</p>' };

const bodyStart = (html: string) => html.slice(html.search(/<body[^>]*>/i)).replace(/<body[^>]*>/i, '');

describe('renderEmail (rich mode)', () => {
  const out = renderEmail({ content: rich(), sender, row: { Name: 'Ravi <b>&</b> Co' } });

  it('substitutes the subject without escaping and the body with escaping', () => {
    expect(out.subject).toBe('Review meeting for Ravi <b>&</b> Co');
    expect(out.html).toContain('Dear Ravi &lt;b&gt;&amp;&lt;/b&gt; Co,');
    expect(out.html).not.toContain('Dear Ravi <b>');
  });
  it('includes the substituted signature', () => {
    expect(out.html).toContain('IQAC Office for Ravi');
  });
  it('omits the signature when includeSignature is false', () => {
    const o = renderEmail({ content: rich({ includeSignature: false }), sender, row: { Name: 'A' } });
    expect(o.html).not.toContain('IQAC Office for');
  });
  it('puts the hidden preheader first in <body>', () => {
    expect(bodyStart(out.html).trimStart().startsWith('<div style="display:none')).toBe(true);
    expect(bodyStart(out.html)).toContain('Agenda for Ravi');
  });
  it('locks light mode (meta + embedded :root style survives inlining)', () => {
    expect(out.html).toMatch(/<meta name="color-scheme" content="light only">/);
    expect(out.html).toMatch(/<style[^>]*>:root\{color-scheme:light only/);
  });
  it('has a text part without the preheader or HTML', () => {
    expect(out.text).toContain('Dear Ravi <b>&</b> Co,');
    expect(out.text).not.toContain('Agenda for');
    expect(out.text).not.toMatch(/<p>|<div/);
  });
  it('reports size and lint; no unresolved placeholders when the row is complete', () => {
    expect(out.sizeBytes).toBe(Buffer.byteLength(out.html));
    expect(out.lint.map(l => l.code)).not.toContain('unresolved_placeholder');
  });
  it('flags unresolved placeholders when a value is missing', () => {
    const o = renderEmail({ content: rich(), sender, row: {} });
    expect(o.lint.find(l => l.code === 'unresolved_placeholder')?.message).toContain('{Name}');
  });
  it('uses a hand-written text part when given', () => {
    const o = renderEmail({ content: rich({ bodyText: 'Hello {Name}' }), sender, row: { Name: 'Asha' } });
    expect(o.text).toBe('Hello Asha');
  });
  it('lists enhanced web fonts chosen by the author and links them once', () => {
    const o = renderEmail({ content: rich({ bodyHtml: '<p style="font-family:\'Playfair Display\', Georgia, serif">Hello everyone at the department, welcome.</p>' }), sender, row: {} });
    expect(o.webFontsUsed).toContain('Playfair Display');
    expect(o.html.match(/family=Playfair\+Display/g)?.length).toBe(1);
  });
  it('passes the unsubscribe URL to the shell', () => {
    const o = renderEmail({ content: rich(), sender, row: { Name: 'A' }, unsubscribeUrl: 'https://mailer.example/u/tok123' });
    expect(o.html).toContain('https://mailer.example/u/tok123');
  });
});

describe('renderEmail (html mode)', () => {
  const doc = `<!DOCTYPE html><html><head><style>.lead{color:#0b5394;font-size:18px} @media only screen and (max-width:600px){.lead{font-size:16px !important}}</style></head><body><p class="lead">Hello {Name}</p><p>Plenty of body text for the department heads to read before the meeting.</p></body></html>`;
  const out = renderEmail({ content: rich({ editorMode: 'html', bodyHtml: doc, preheader: 'Pre' }), sender, row: { Name: 'Asha' } });

  it('inlines CSS and keeps media queries', () => {
    expect(out.html).toMatch(/<p class="lead" style="color: #0b5394; font-size: 18px;">Hello Asha<\/p>/);
    expect(out.html).toContain('@media only screen and (max-width:600px)');
    expect(out.html).not.toContain('.lead{color:#0b5394');
  });
  it('does not wrap the document in a theme shell or add the signature', () => {
    expect(out.html).not.toContain('IQAC Office for');
  });
  it('wraps a bare fragment into a document', () => {
    const o = renderEmail({ content: rich({ editorMode: 'html', bodyHtml: '<p>Just a fragment with enough words in it.</p>' }), sender, row: {} });
    expect(o.html).toMatch(/<html[^>]*>[\s\S]*<body[^>]*>/);
  });
});

describe('htmlToText', () => {
  it('keeps link URLs, skips images and does not uppercase headings', () => {
    const t = htmlToText('<h1>Annual Review</h1><p>See <a href="https://kiit.ac.in/iqac">the page</a>.</p><img src="x.png" alt="logo">');
    expect(t).toContain('Annual Review');
    expect(t).not.toContain('ANNUAL REVIEW');
    expect(t).toContain('[https://kiit.ac.in/iqac]');
    expect(t).not.toContain('logo');
  });
});

describe('hosted image → CID conversion', () => {
  const id = '0b8f3c1e-1111-4222-8333-944455556666';
  const html = `<img src="/api/uploads/${id}/content" alt="a"><img src='http://localhost:5173/api/uploads/${id}/content' alt="b"><img src="https://x.public.blob.vercel-storage.com/y.png" alt="c">`;
  it('finds app-hosted image ids', () => {
    process.env.APP_BASE_URL = 'http://localhost:5173';
    expect(hostedImageIds(html)).toEqual([id]);
  });
  it('rewrites to cid: and returns inline attachments', () => {
    process.env.APP_BASE_URL = 'http://localhost:5173';
    const r = inlineHostedImages(html, [{ id, filename: 'a.png', contentType: 'image/png', content: Buffer.from('png') }]);
    expect(r.html).toContain(`src="cid:${cidFor(id)}"`);
    expect(r.html).toContain(`src='cid:${cidFor(id)}'`);
    expect(r.html).toContain('https://x.public.blob.vercel-storage.com/y.png');
    expect(r.attachments).toHaveLength(1);
    expect(r.attachments[0]).toMatchObject({ cid: `${id}@kiit-mailer`, contentDisposition: 'inline' });
  });
  it('leaves unknown images alone', () => {
    const r = inlineHostedImages(html, []);
    expect(r.html).toBe(html);
    expect(r.attachments).toEqual([]);
  });
});
