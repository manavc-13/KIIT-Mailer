import { describe, expect, it } from 'vitest';
import { detectFeatures, hasOutlookConditionals } from '../src/preview/features';
import { CANIEMAIL, compatibilityReport, parseSupport, supportFor } from '../src/preview/caniemail';
import { FEATURE_SLUGS } from '../src/preview/caniemail-spec';
import { injectPreheader } from '../shared/text';

const SAMPLE = `<!DOCTYPE html><html><head>
<link href="https://fonts.googleapis.com/css2?family=Lato" rel="stylesheet">
<style>@media only screen and (max-width:600px){.wrap{width:100%!important}} a:hover{color:red}</style>
</head><body>
<table role="presentation" width="600" style="max-width:600px;border-radius:8px;background-image:url(https://x/bg.png)">
<tr><td style="padding:16px;font-family:Lato, 'Helvetica Neue', Arial, sans-serif">
<img src="https://x/logo.webp" alt="" style="display:block;margin:0 auto">
<div style="display:flex;gap:8px;padding:4px"><span>a</span></div>
<svg width="10" height="10"></svg>
</td></tr></table>
<!--[if mso]><table><tr><td><![endif]-->
</body></html>`;

describe('feature detection', () => {
  const slugs = detectFeatures(SAMPLE).map(f => f.slug);

  it('finds the features used', () => {
    for (const s of [
      'html-style', 'css-at-media', 'css-pseudo-class-hover', 'css-at-font-face', 'css-max-width', 'css-border-radius',
      'css-background-image', 'css-display-flex', 'css-gap', 'css-margin', 'css-padding', 'html-svg', 'image-webp',
    ]) expect(slugs, s).toContain(s);
  });

  it('does not report features that are absent', () => {
    for (const s of ['css-display-grid', 'css-position', 'html-video', 'css-variables', 'css-box-shadow', 'image-base64']) {
      expect(slugs, s).not.toContain(s);
    }
  });

  it('only flags padding on non-table-cell elements', () => {
    const f = detectFeatures(SAMPLE).find(x => x.slug === 'css-padding');
    expect(f?.evidence).toMatch(/<div>/);
    expect(detectFeatures('<td style="padding:10px">x</td>').map(x => x.slug)).not.toContain('css-padding');
  });

  it('ignores harmless zero values and normal margins', () => {
    const slugs2 = detectFeatures('<p style="margin:0 0 12px;border-radius:0;max-width:none">x</p>').map(f => f.slug);
    expect(slugs2).toEqual([]);
  });

  it("ignores the preheader's hiding tricks", () => {
    const html = injectPreheader('<html><body><p>Hi</p></body></html>', 'Preview text');
    expect(detectFeatures(html)).toEqual([]);
  });

  it('detects CSS variables, grid, position and data URIs', () => {
    const s = detectFeatures(`<div style="--x:1;display:grid;position:absolute"><img src="data:image/png;base64,AA"></div>`).map(f => f.slug);
    expect(s).toEqual(expect.arrayContaining(['css-variables', 'css-display-grid', 'css-position', 'image-base64']));
  });

  it('detects Outlook conditional comments', () => {
    expect(hasOutlookConditionals(SAMPLE)).toBe(true);
    expect(hasOutlookConditionals('<p>x</p>')).toBe(false);
  });
});

describe('caniemail lookup', () => {
  it('vendored data covers every detectable feature and is small', () => {
    for (const slug of FEATURE_SLUGS) expect(CANIEMAIL.features[slug], slug).toBeTruthy();
    expect(JSON.stringify(CANIEMAIL).length).toBeLessThan(60 * 1024);
  });

  it('parses support values with notes', () => {
    expect(parseSupport('a #1 #2', { 1: 'one', 2: 'two' })).toEqual({ level: 'a', notes: ['one', 'two'] });
    expect(parseSupport('y')).toEqual({ level: 'y', notes: [] });
    expect(parseSupport(undefined).level).toBe('u');
  });

  it('knows the well-established cases', () => {
    expect(supportFor('css-border-radius', 'outlook-windows').level).toBe('n');
    expect(supportFor('css-border-radius', 'gmail-desktop').level).toBe('y');
    expect(supportFor('css-display-flex', 'outlook-windows').level).toBe('n');
    expect(supportFor('css-at-font-face', 'gmail-desktop').level).toBe('n');
    expect(supportFor('css-at-font-face', 'apple-mail').level).toBe('y');
    expect(supportFor('css-at-media', 'outlook-windows').level).toBe('n');
  });

  it('a target with several platforms takes the worst one', () => {
    const db = {
      ...CANIEMAIL,
      features: { 'css-x': { t: 'x', u: '', c: 'css', s: { 'gmail-ios': 'y', 'gmail-android': 'a #1' }, n: { 1: 'Android only' } } },
    };
    expect(supportFor('css-x', 'gmail-mobile', db)).toEqual({ level: 'a', notes: ['Android only'] });
  });

  it('builds a report with problem rows first', () => {
    const rows = compatibilityReport(SAMPLE);
    expect(rows.length).toBeGreaterThan(5);
    const radius = rows.find(r => r.slug === 'css-border-radius')!;
    expect(radius.problems).toContain('outlook-windows');
    expect(radius.title).toBe('border-radius');
    expect(rows[0].problems.length).toBeGreaterThan(0);
    const last = rows[rows.length - 1];
    expect(rows[0].problems.length).toBeGreaterThanOrEqual(last.problems.length);
  });
});
