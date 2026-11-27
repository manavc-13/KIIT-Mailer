import { describe, expect, it } from 'vitest';
import {
  backgroundColorOf, darkModeSimulation, gmailClip, gmailTransform, htmlForClient, mapInlineStyles,
  outlookTransform, prepareFrameHtml, removeAtRules, splitDeclarations, stripWebFonts, utf8Length,
} from '../src/preview/transforms';
import { GMAIL_CLIP_BYTES } from '../shared/text';

const doc = (head: string, body: string) => `<!DOCTYPE html><html><head>${head}</head><body>${body}</body></html>`;

describe('CSS helpers', () => {
  it('splits declarations without breaking data URLs or quoted values', () => {
    expect(splitDeclarations("color:red; background:url(data:image/png;base64,AAA=) no-repeat; font-family:'A;B', serif"))
      .toEqual(['color:red', 'background:url(data:image/png;base64,AAA=) no-repeat', "font-family:'A;B', serif"]);
  });

  it('decodes &quot; inside style attributes before splitting', () => {
    const html = '<p style="font-family:&quot;Open Sans&quot;, Arial;color:#111">x</p>';
    const out = mapInlineStyles(html, d => d);
    expect(out).toBe(`<p style="font-family:'Open Sans', Arial;color:#111">x</p>`);
  });

  it('removes at-rules with nested braces', () => {
    const css = 'p{color:red}@media (max-width:600px){.a{width:100%}.b{display:none}}@import url(x.css);h1{margin:0}';
    expect(removeAtRules(css, ['media', 'import'])).toBe('p{color:red}h1{margin:0}');
  });

  it('salvages the colour from a background shorthand', () => {
    expect(backgroundColorOf("#0c1a28 url('hero.jpg') no-repeat center / cover")).toBe('#0c1a28');
    expect(backgroundColorOf('url(a.png) no-repeat center')).toBeNull();
    expect(backgroundColorOf('linear-gradient(90deg, #fff, #000) rgba(0, 0, 0, 0.5)')).toBe('rgba(0, 0, 0, 0.5)');
    expect(backgroundColorOf('url(a.png) navy')).toBe('navy');
  });
});

describe('Outlook (Windows) approximation', () => {
  const sample = doc(
    `<link href="https://fonts.googleapis.com/css2?family=Lato&display=swap" rel="stylesheet">
     <style>@font-face{font-family:X;src:url(x.woff)} .btn{border-radius:6px;color:#fff} @media (max-width:600px){.wrap{width:100%!important}}</style>`,
    `<div class="wrap" style="max-width:600px;margin:0 auto;border-radius:8px;box-shadow:0 1px 3px #000;background-image:url(bg.png);color:#222">
       <td background="bg.png" style="background:#0c1a28 url(bg.png) no-repeat;padding:12px">x</td>
       <div style="display:flex;gap:8px"><span style="font-family:Lato, 'Helvetica Neue', Arial, sans-serif">Hi</span></div>
       <!--[if mso]><table width="600"><tr><td><![endif]-->
       <!--[if mso]><v:roundrect arcsize="10%"><center>Btn</center></v:roundrect><![endif]-->
     </div>`,
  );
  const out = outlookTransform(sample);

  it('drops border-radius, box-shadow, max-width and background images', () => {
    expect(out).not.toMatch(/border-radius/);
    expect(out).not.toMatch(/box-shadow/);
    expect(out).not.toMatch(/max-width/);
    expect(out).not.toMatch(/background-image/);
    expect(out).not.toMatch(/url\(bg\.png\)/);
    expect(out).not.toMatch(/background="bg\.png"/);
  });

  it('keeps supported declarations and the background colour', () => {
    expect(out).toMatch(/margin:0 auto/);
    expect(out).toMatch(/color:#222/);
    expect(out).toMatch(/background-color:#0c1a28/);
    expect(out).toMatch(/padding:12px/);
    expect(out).toMatch(/\.btn\{color:#fff\}/);
  });

  it('drops media queries, @font-face, web font links and flex', () => {
    expect(out).not.toMatch(/@media/);
    expect(out).not.toMatch(/@font-face/);
    expect(out).not.toMatch(/fonts\.googleapis/);
    expect(out).not.toMatch(/display:flex/);
    expect(out).toMatch(/font-family:'Helvetica Neue', Arial, sans-serif/);
  });

  it('reveals non-VML Outlook conditionals but keeps VML hidden', () => {
    expect(out).toMatch(/<table width="600"><tr><td>/);
    expect(out).toMatch(/<!--\[if mso\]><v:roundrect/);
  });

  it('leaves [if !mso] blocks alone', () => {
    const html = '<!--[if !mso]><!--><a href="#">x</a><!--<![endif]-->';
    expect(outlookTransform(html)).toBe(html);
  });
});

describe('Gmail', () => {
  it('replaces web fonts with their fallback and removes font links', () => {
    const html = doc('<link href="https://fonts.googleapis.com/css2?family=Playfair+Display" rel="stylesheet">',
      `<h1 style="font-family:'Playfair Display', Georgia, 'Times New Roman', serif">T</h1>`);
    const out = stripWebFonts(html);
    expect(out).not.toMatch(/Playfair/);
    expect(out).toMatch(/font-family:Georgia/);
  });

  it('does not clip at or under the limit', () => {
    const html = '<p>' + 'a'.repeat(1000) + '</p>';
    expect(gmailClip(html, 1007, 2000)).toEqual({ html, clipped: false });
  });

  it('clips at the byte limit, never inside a tag', () => {
    const html = '<p>' + 'a'.repeat(95) + '<a href="https://example.com/long">link</a></p>';
    const r = gmailClip(html, utf8Length(html), 110);
    expect(r.clipped).toBe(true);
    expect(r.html).toBe('<p>' + 'a'.repeat(95)); // the cut at byte 110 fell inside <a ...>
  });

  it('counts multi-byte characters as UTF-8 bytes', () => {
    const html = 'é'.repeat(100); // 200 bytes
    const r = gmailClip(html, utf8Length(html), 51);
    expect(r.html).toBe('é'.repeat(25));
  });

  it('does not leave half an HTML entity', () => {
    const html = 'abc&nbsp;def';
    expect(gmailClip(html, 99, 6).html).toBe('abc');
  });

  it('clips real-size messages at GMAIL_CLIP_BYTES', () => {
    const big = doc('', '<p>' + 'x'.repeat(GMAIL_CLIP_BYTES + 5000) + '</p>');
    const r = gmailTransform(big);
    expect(r.clipped).toBe(true);
    expect(utf8Length(r.html)).toBeLessThanOrEqual(GMAIL_CLIP_BYTES);
    expect(utf8Length(r.html)).toBeGreaterThan(GMAIL_CLIP_BYTES - 16);
  });
});

describe('frame preparation', () => {
  it('removes scripts and adds a base target', () => {
    const out = prepareFrameHtml(doc('<script>alert(1)</script>', '<p>x</p><script src="x.js"></script>'));
    expect(out).not.toMatch(/<script/i);
    expect(out).toMatch(/<base target="_blank">/);
  });

  it('dark simulation appends an invert filter after the email styles', () => {
    const out = darkModeSimulation(doc('<style>p{color:red}</style>', '<p>x</p>'));
    expect(out.indexOf('data-kiit-dark')).toBeGreaterThan(out.indexOf('p{color:red}'));
    expect(out).toMatch(/invert\(1\)/);
  });

  it('htmlForClient keeps web fonts for Apple Mail only', () => {
    const html = doc('', `<p style="font-family:Lato, 'Helvetica Neue', Arial, sans-serif">x</p>`);
    expect(htmlForClient('apple-iphone', html).html).toMatch(/font-family:Lato/);
    expect(htmlForClient('gmail-mobile', html).html).not.toMatch(/Lato/);
    expect(htmlForClient('outlook-windows', html).html).not.toMatch(/Lato/);
  });
});
