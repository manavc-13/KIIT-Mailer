import { describe, expect, it } from 'vitest';
import juice from 'juice';
import type { ThemeId } from '../shared/types.js';
import { BLOCK_CLASSES } from '../shared/blocks.js';
import {
  THEMES, DEFAULT_BRAND, themeCss, renderShell, starterSignatureHtml, themeFonts, getTheme, mix, msoCss, mobileCss,
} from '../shared/themes/index.js';

const IDS: ThemeId[] = ['heritage', 'meridian', 'letterhead', 'slate', 'plain'];

const BODY =
  '<p>Hello <strong>world</strong></p>' +
  '<table role="presentation" class="kb-callout" data-variant="warn" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-callout-cell"><p>Careful</p></td></tr></table>' +
  '<table role="presentation" class="kb-details" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-details-label">Date</td><td class="kb-details-value">12 Dec 2026</td></tr></table>' +
  '<table role="presentation" class="kb-button" data-variant="primary" align="center" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-button-cell"><a class="kb-button-link" href="https://kiit.ac.in" target="_blank">Go</a></td></tr></table>' +
  '<table role="presentation" class="kb-divider" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-divider-cell">&nbsp;</td></tr></table>';
const SIG = starterSignatureHtml({ name: 'Test Sender', title: 'Coordinator', email: 'sender@kiit.ac.in' });

describe('THEMES metadata', () => {
  it('lists every ThemeId once with label, description and swatch', () => {
    expect(THEMES.map(t => t.id).sort()).toEqual([...IDS].sort());
    for (const t of THEMES) {
      expect(t.label).toBeTruthy();
      expect(t.description.length).toBeGreaterThan(20);
      expect(t.swatch.length).toBeGreaterThanOrEqual(3);
      expect(t.swatch.length).toBeLessThanOrEqual(4);
      for (const c of t.swatch) expect(c).toMatch(/^#[0-9A-F]{6}$/i);
    }
  });

  it('ports the skill tokens 1:1 (spot checks)', () => {
    expect(getTheme('heritage').colors.header_bg).toBe('#0C1A28');
    expect(getTheme('heritage').colors.accent).toBe('#C5973A');
    expect(getTheme('meridian').colors.accent).toBe('#17D059');
    expect(getTheme('letterhead').colors.accent).toBe('#7A1F2B');
    expect(getTheme('slate').colors.button_bg).toBe('#2550D6');
    expect(getTheme('letterhead').header_style).toBe('light');
    expect(getTheme('plain').header_style).toBe('plain');
  });

  it('mix() matches build_email.py', () => {
    // python: mix('#C5973A', '#0C1A28', 0.35)
    expect(mix('#C5973A', '#0C1A28', 0.35)).toBe('#4D462E');
    expect(mix('#17D059', '#081710', 0.2)).toBe('#0B3C1F');
    expect(mix('#000000', '#FFFFFF', 0.5)).toBe('#808080');
  });

  it('font stacks end in a web-safe face', () => {
    for (const id of IDS) {
      const f = themeFonts(id);
      for (const stack of [f.display, f.body]) expect(stack).toMatch(/(serif|sans-serif|monospace)$/);
    }
  });

  it('DEFAULT_BRAND comes from brand.json', () => {
    expect(DEFAULT_BRAND.logoUrl).toMatch(/^https:\/\/.+kiit-logo\.png$/);
    expect(DEFAULT_BRAND.title).toBe('Kalinga Institute of Industrial Technology');
    expect(DEFAULT_BRAND.logoWidth).toBe(88);
  });
});

describe.each(IDS)('renderShell(%s)', id => {
  const html = renderShell({ theme: id, bodyHtml: BODY, signatureHtml: SIG, subject: 'Offer & <Terms> "2026"' });

  it('renders a full email-safe document', () => {
    expect(html.startsWith('<!DOCTYPE html>')).toBe(true);
    expect(html).toContain('<meta charset="UTF-8">');
    expect(html).toContain('name="viewport"');
    expect(html).toContain('x-apple-disable-message-reformatting');
    expect(html).toContain('name="format-detection"');
    expect(html).toContain('name="color-scheme"');
    expect(html).toContain('role="presentation"');
    expect(html).toContain('<!--[if mso]>');
    expect(html).toContain('max-width:600px');
    expect(html).toMatch(/<\/html>\s*$/);
  });

  it('contains kb-body, the body, the signature and the escaped subject', () => {
    expect(html).toMatch(/<td class="kb-body[^"]*"/);
    expect(html).toContain(BODY);
    expect(html).toContain('class="kb-signature"');
    expect(html).toContain('Test Sender');
    expect(html).toContain('<title>Offer &amp; &lt;Terms&gt; &quot;2026&quot;</title>');
  });

  it('shows the logo with width and alt, and brand footer', () => {
    expect(html).toContain(`src="${DEFAULT_BRAND.logoUrl}"`);
    expect(html).toMatch(/<img[^>]+alt="KIIT"[^>]+width="\d+"/);
    expect(html).toContain('KIIT Deemed to be University');
    expect(html).toContain(DEFAULT_BRAND.disclaimer);
  });

  it('only shows Unsubscribe when a URL is given', () => {
    expect(html).not.toContain('Unsubscribe');
    const withUnsub = renderShell({ theme: id, bodyHtml: BODY, signatureHtml: null, subject: 's', unsubscribeUrl: 'https://x.test/u/a?b=1&c=2' });
    expect(withUnsub).toContain('>Unsubscribe</a>');
    expect(withUnsub).toContain('href="https://x.test/u/a?b=1&amp;c=2"');
  });

  it('omits the signature when null', () => {
    const noSig = renderShell({ theme: id, bodyHtml: BODY, signatureHtml: null, subject: 's' });
    expect(noSig).not.toContain('class="kb-signature"');
  });

  it('loads web fonts only for themes that use them', () => {
    const url = getTheme(id).fonts.webfont_url;
    if (url) expect(html).toContain('fonts.googleapis.com');
    else expect(html).not.toContain('fonts.googleapis.com');
  });

  it('contains no CSS custom properties', () => {
    expect(html).not.toContain('var(--');
  });

  it('respects brand overrides (empty logo drops it)', () => {
    const h = renderShell({ theme: id, bodyHtml: '', signatureHtml: null, subject: 's', brand: { logoUrl: '', title: 'Custom Title' } });
    expect(h).not.toContain('<img');
    expect(h).toContain('Custom Title');
  });

  it('survives juice: body inlined, media queries kept', () => {
    const out = juice(html, { preserveMediaQueries: true, removeStyleTags: true, applyAttributesTableElements: true, applyWidthAttributes: true });
    const t = getTheme(id);
    expect(out).toContain('@media');
    expect(out).toMatch(/class="kb-callout-cell"[^>]*style="[^"]*background-color: ?#/);
    expect(out).toContain(t.colors.warn_bg);
    expect(out).toMatch(/class="kb-button-link"[^>]*style="[^"]*text-decoration: ?none/);
    expect(out).toMatch(/class="kb-divider-cell"[^>]*height="1"/);
    expect(out).not.toContain('var(--');
  });
});

describe('themeCss', () => {
  const classes = Object.values(BLOCK_CLASSES).filter(c => c !== BLOCK_CLASSES.body);

  it.each(IDS)('%s covers every block class and the signature', id => {
    const css = themeCss(id);
    for (const c of classes) expect(css).toContain(`.${c}`);
    expect(css).toContain('.kb-signature');
    expect(css).toContain('[data-variant="warn"]');
    expect(css).toContain('[data-variant="success"]');
    expect(css).toContain('[data-variant="secondary"]');
    for (const el of ['p', 'h1', 'h2', 'h3', 'a', 'ul', 'li', 'blockquote', 'table', 'td', 'hr', 'img']) {
      expect(css).toMatch(new RegExp(`\\.kb-body ${el}[,{]`));
    }
  });

  it.each(IDS)('%s has no CSS variables or unsupported constructs', id => {
    for (const css of [themeCss(id), mobileCss(id), msoCss(id)]) {
      expect(css).not.toContain('var(--');
      expect(css).not.toMatch(/rgba\(|gradient|box-shadow|:has\(|display:\s*(flex|grid)/);
    }
  });

  it('scopes every rule under the given scope', () => {
    const css = themeCss('slate', '.editor .kb-body');
    for (const line of css.split('\n')) {
      for (const sel of line.slice(0, line.indexOf('{')).split(',')) expect(sel.trim().startsWith('.editor .kb-body')).toBe(true);
    }
  });

  it('heritage divider is gold', () => {
    expect(themeCss('heritage')).toMatch(/\.kb-divider-cell\{[^}]*background-color:#C5973A/);
  });
});

describe('starterSignatureHtml', () => {
  it('builds the standard IQAC signature', () => {
    const s = starterSignatureHtml({ name: 'A & B', title: 'Intern', email: 'a@kiit.ac.in', phone: '+91 98765 43210' });
    expect(s).toContain('Warm regards,');
    expect(s).toContain('A &amp; B');
    expect(s).toContain('Intern');
    expect(s).toContain('Internal Quality Assurance Cell (IQAC)');
    expect(s).toContain('KIIT Deemed to be University');
    expect(s).toContain('href="mailto:a@kiit.ac.in"');
    expect(s).toContain('href="tel:+919876543210"');
    expect(s).not.toMatch(/style=/);
  });

  it('uses department when given and skips optional parts', () => {
    const s = starterSignatureHtml({ name: 'X', email: 'x@kiit.ac.in', department: 'School of Law' });
    expect(s).toContain('School of Law');
    expect(s).not.toContain('Internal Quality Assurance Cell (IQAC)');
    expect(s).not.toContain('kb-signature-title');
    expect(s).not.toContain('tel:');
  });
});
