import { describe, it, expect } from 'vitest';
import { lintEmail } from '../server/lib/lint.js';

const doc = (body: string) => `<!DOCTYPE html><html><head><style>@media (max-width:600px){.x{width:100%}}</style></head><body><div style="display:none;max-height:0">Preheader {NotAPlaceholderInHidden}</div>${body}</body></html>`;
const LONG_TEXT = '<p>' + 'The Internal Quality Assurance Cell invites all heads of departments to the annual review meeting. '.repeat(4) + '</p>';
const codes = (subject: string, body: string, text: string | null = 'Plain text part') =>
  lintEmail({ subject, html: doc(body), text }).map(i => i.code);

describe('lintEmail', () => {
  it('a clean institutional email has no issues', () => {
    expect(lintEmail({ subject: 'IQAC review meeting on 12 December', html: doc(LONG_TEXT + '<a href="https://kiit.ac.in/iqac">IQAC page</a>'), text: 'x' })).toEqual([]);
  });

  it('spam words in subject', () => {
    expect(codes('Claim your FREE prize now', LONG_TEXT)).toContain('spam_word');
    expect(codes('Annual review meeting', LONG_TEXT)).not.toContain('spam_word');
  });
  it('spam phrases in body (but routine words like "offer" are fine in the body)', () => {
    expect(codes('Meeting', LONG_TEXT + '<p>Click here to act now!</p>')).toContain('spam_word_body');
    expect(codes('Meeting', LONG_TEXT + '<p>Please find your internship offer letter attached.</p>')).not.toContain('spam_word_body');
  });
  it('ALL-CAPS subject', () => {
    expect(codes('IMPORTANT NOTICE FOR ALL STAFF', LONG_TEXT)).toContain('caps_subject');
    expect(codes('IQAC: NAAC review notice', LONG_TEXT)).not.toContain('caps_subject');
  });
  it('!!! or $$ in subject', () => {
    expect(codes('Meeting today!!!', LONG_TEXT)).toContain('subject_punctuation');
    expect(codes('Win $$ now', LONG_TEXT)).toContain('subject_punctuation');
    expect(codes('Meeting today!', LONG_TEXT)).not.toContain('subject_punctuation');
  });
  it('subject longer than 78 characters', () => {
    expect(codes('A'.repeat(10) + ' b'.repeat(40), LONG_TEXT)).toContain('subject_long');
    expect(codes('Short subject', LONG_TEXT)).not.toContain('subject_long');
  });
  it('empty subject', () => {
    const issues = lintEmail({ subject: '  ', html: doc(LONG_TEXT), text: 'x' });
    expect(issues.find(i => i.code === 'empty_subject')?.level).toBe('error');
    expect(codes('Subject', LONG_TEXT)).not.toContain('empty_subject');
  });
  it('data: image URIs are errors', () => {
    const issues = lintEmail({ subject: 'Hi', html: doc(LONG_TEXT + '<img alt="x" src="data:image/png;base64,AAAA">'), text: 'x' });
    expect(issues.find(i => i.code === 'data_uri_image')?.level).toBe('error');
    expect(codes('Hi', LONG_TEXT + '<img alt="x" src="https://kiit.ac.in/logo.png">')).not.toContain('data_uri_image');
  });
  it('images missing alt', () => {
    expect(codes('Hi', LONG_TEXT + '<img src="https://kiit.ac.in/a.png">')).toContain('missing_alt');
    expect(codes('Hi', LONG_TEXT + '<img src="https://kiit.ac.in/a.png" alt="">')).not.toContain('missing_alt');
  });
  it('image-only and low text ratio', () => {
    expect(codes('Hi', '<img alt="poster" src="https://kiit.ac.in/poster.png">')).toContain('image_only');
    expect(codes('Hi', '<p>See the posters below for the full programme.</p>' + '<img alt="p" src="https://kiit.ac.in/p.png">'.repeat(3))).toContain('low_text_ratio');
    const ok = codes('Hi', LONG_TEXT + '<img alt="p" src="https://kiit.ac.in/p.png">');
    expect(ok).not.toContain('image_only');
    expect(ok).not.toContain('low_text_ratio');
  });
  it('http:// links', () => {
    expect(codes('Hi', LONG_TEXT + '<a href="http://kiit.ac.in">site</a>')).toContain('http_link');
    expect(codes('Hi', LONG_TEXT + '<a href="https://kiit.ac.in">site</a>')).not.toContain('http_link');
  });
  it('URL shorteners', () => {
    expect(codes('Hi', LONG_TEXT + '<a href="https://bit.ly/abc">form</a>')).toContain('shortener');
    expect(codes('Hi', LONG_TEXT + '<a href="https://forms.gle/abc">form</a>')).not.toContain('shortener');
  });
  it('link text domain that differs from the href domain is an error', () => {
    const issues = lintEmail({ subject: 'Hi', html: doc(LONG_TEXT + '<a href="https://evil.example.com/login">www.kiit.ac.in</a>'), text: 'x' });
    expect(issues.find(i => i.code === 'link_mismatch')?.level).toBe('error');
    expect(codes('Hi', LONG_TEXT + '<a href="https://admissions.kiit.ac.in/apply">kiit.ac.in/apply</a>')).not.toContain('link_mismatch');
    expect(codes('Hi', LONG_TEXT + '<a href="https://evil.example.com">Apply here</a>')).not.toContain('link_mismatch');
  });
  it('HTML over the Gmail clip size', () => {
    const big = LONG_TEXT + `<p>${'x'.repeat(110 * 1024)}</p>`;
    expect(codes('Hi', big)).toContain('gmail_clip');
    expect(codes('Hi', LONG_TEXT)).not.toContain('gmail_clip');
  });
  it('unresolved placeholders are errors (CSS braces and hidden preheader do not count)', () => {
    const issues = lintEmail({ subject: 'Hello {Name}', html: doc(LONG_TEXT + '<p>Dear {Name}, your dept is {Department}.</p>'), text: 'x' });
    const u = issues.find(i => i.code === 'unresolved_placeholder');
    expect(u?.level).toBe('error');
    expect(u?.message).toContain('{Name}');
    expect(u?.message).toContain('{Department}');
    expect(codes('Hello Ravi', LONG_TEXT)).not.toContain('unresolved_placeholder');
  });
  it('missing text part is info', () => {
    const issues = lintEmail({ subject: 'Hi', html: doc(LONG_TEXT), text: '' });
    expect(issues.find(i => i.code === 'no_text')?.level).toBe('info');
    expect(codes('Hi', LONG_TEXT, 'text')).not.toContain('no_text');
  });
  it('JavaScript, forms and iframes are errors', () => {
    expect(codes('Hi', LONG_TEXT + '<script>alert(1)</script>')).toContain('javascript');
    expect(codes('Hi', LONG_TEXT + '<a href="javascript:alert(1)">x</a>')).toContain('javascript');
    expect(codes('Hi', LONG_TEXT + '<p onclick="x()">x</p>')).toContain('javascript');
    expect(codes('Hi', LONG_TEXT + '<form action="https://x"><input></form>')).toContain('form');
    expect(codes('Hi', LONG_TEXT + '<iframe src="https://x"></iframe>')).toContain('iframe');
    const clean = codes('Hi', LONG_TEXT);
    expect(clean).not.toContain('javascript');
    expect(clean).not.toContain('form');
    expect(clean).not.toContain('iframe');
  });
});
