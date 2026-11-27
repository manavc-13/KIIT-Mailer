import { describe, it, expect, beforeAll } from 'vitest';
import crypto from 'node:crypto';
import { classifyError, buildMailOptions, maxRawAttachmentBytes, senderCredentials } from '../server/lib/mailer.js';
import { makeUnsubscribeToken, parseUnsubscribeToken, unsubscribeUrl, confirmPage, donePage } from '../server/lib/unsubscribe.js';
import { seal } from '../server/lib/crypto.js';
import { isBlockedIp, assertFetchableUrl } from '../server/lib/send/ssrf.js';

beforeAll(() => {
  process.env.APP_ENCRYPTION_KEY ||= crypto.randomBytes(32).toString('base64');
  process.env.APP_BASE_URL = 'https://mailer.example';
});

describe('classifyError', () => {
  it('auth failures (535) are not retryable', () => {
    const r = classifyError({ responseCode: 535, response: '535-5.7.8 Username and Password not accepted' });
    expect(r.kind).toBe('auth');
    expect(r.retryable).toBe(false);
    expect(r.message).toContain('App Password rejected');
  });
  it('daily quota', () => {
    expect(classifyError({ response: '550-5.4.5 Daily user sending limit exceeded' }).kind).toBe('quota');
  });
  it('throttling and network errors are retryable', () => {
    expect(classifyError({ responseCode: 421, response: '421 4.7.0 Temporary System Problem' })).toMatchObject({ kind: 'throttle', retryable: true });
    expect(classifyError({ code: 'ETIMEDOUT', message: 'Connection timeout' })).toMatchObject({ kind: 'network', retryable: true });
  });
  it('size', () => {
    expect(classifyError({ response: '552 5.3.4 Message size exceeds fixed limit' }).kind).toBe('size');
  });
  it('unknown errors keep the raw message', () => {
    expect(classifyError(new Error('weird'))).toMatchObject({ kind: 'other', message: 'weird', retryable: false });
  });
});

describe('buildMailOptions', () => {
  const sender = { email: 'iqac@kiit.ac.in', displayName: 'IQAC, "KIIT"', replyTo: 'office@kiit.ac.in' };
  it('always sends from the account, with both parts, Message-ID and no X-Mailer', () => {
    const o = buildMailOptions({ sender, to: 'a@x.com', subject: 'S', html: '<p>x</p>', text: 'x' });
    expect(o.from).toEqual({ name: 'IQAC, "KIIT"', address: 'iqac@kiit.ac.in' });
    expect(o.replyTo).toBe('office@kiit.ac.in');
    expect(o.html).toBe('<p>x</p>');
    expect(o.text).toBe('x');
    expect(String(o.messageId)).toMatch(/^<[0-9a-f-]{36}@kiit\.ac\.in>$/);
    expect(o.date).toBeInstanceOf(Date);
    expect((o as { xMailer?: unknown }).xMailer).toBe(false);
    expect(o.headers).toEqual({});
  });
  it('adds List-Unsubscribe with one-click when enabled', () => {
    const o = buildMailOptions({ sender, to: 'a@x.com', subject: 'S', html: 'h', text: 't', unsubscribe: true, unsubscribeUrl: 'https://mailer.example/u/abc' });
    expect(o.headers).toEqual({
      'List-Unsubscribe': '<mailto:iqac@kiit.ac.in?subject=unsubscribe>, <https://mailer.example/u/abc>',
      'List-Unsubscribe-Post': 'List-Unsubscribe=One-Click',
    });
  });
  it('mailto-only List-Unsubscribe without an https URL (BCC broadcast)', () => {
    const o = buildMailOptions({ sender, to: 'iqac@kiit.ac.in', subject: 'S', html: 'h', text: 't', unsubscribe: true, unsubscribeUrl: 'mailto:iqac@kiit.ac.in?subject=unsubscribe' });
    expect(o.headers).toEqual({ 'List-Unsubscribe': '<mailto:iqac@kiit.ac.in?subject=unsubscribe>' });
  });
  it('attachment budget is about 18 MB', () => {
    expect(maxRawAttachmentBytes() / 1024 / 1024).toBeGreaterThan(18);
    expect(maxRawAttachmentBytes() / 1024 / 1024).toBeLessThan(19);
  });
  it('decrypts credentials with the sender id as AAD', () => {
    const id = crypto.randomUUID();
    const s = seal('abcd efgh ijkl mnop', id);
    const row = { id, email: 'iqac@kiit.ac.in', appPasswordEnc: s.enc, appPasswordIv: s.iv, appPasswordTag: s.tag, keyVersion: s.keyVersion };
    expect(senderCredentials(row)).toEqual({ user: 'iqac@kiit.ac.in', pass: 'abcd efgh ijkl mnop' });
    expect(() => senderCredentials({ ...row, id: crypto.randomUUID() })).toThrow(/could not be decrypted/);
  });
});

describe('unsubscribe tokens', () => {
  it('round-trips and lower-cases the address', () => {
    const t = makeUnsubscribeToken('Ravi@Gmail.com', 'sender-1', 'job-1');
    expect(parseUnsubscribeToken(t)).toEqual({ e: 'ravi@gmail.com', s: 'sender-1', j: 'job-1' });
    expect(unsubscribeUrl('a@b.com', 's')).toMatch(/^https:\/\/mailer\.example\/u\/[\w-]+\.[\w-]+$/);
  });
  it('rejects tampered and malformed tokens', () => {
    const t = makeUnsubscribeToken('a@b.com', 's');
    const [body, mac] = t.split('.');
    const forged = Buffer.from(JSON.stringify({ e: 'victim@b.com', s: 's' })).toString('base64url');
    expect(parseUnsubscribeToken(`${forged}.${mac}`)).toBeNull();
    // Always change the first MAC character (the key is random, so the MAC may already start with 'x').
    expect(parseUnsubscribeToken(`${body}.${mac[0] === 'x' ? 'y' : 'x'}${mac.slice(1)}`)).toBeNull();
    expect(parseUnsubscribeToken('garbage')).toBeNull();
    expect(parseUnsubscribeToken('')).toBeNull();
  });
  it('pages escape the address', () => {
    expect(confirmPage('<x>@b.com')).toContain('&lt;x&gt;@b.com');
    expect(confirmPage('a@b.com')).toContain('Confirm unsubscribe');
    expect(donePage('a@b.com')).toContain('You have been unsubscribed');
  });
});

describe('SSRF guard', () => {
  it('blocks private, loopback, link-local and mapped addresses', () => {
    for (const ip of ['127.0.0.1', '10.1.2.3', '172.16.0.1', '192.168.1.1', '169.254.169.254', '100.64.0.1', '0.0.0.0', '::1', '::', 'fe80::1', 'fd00::1', '::ffff:127.0.0.1', '::ffff:7f00:1', '64:ff9b::a00:1']) {
      expect(isBlockedIp(ip), ip).toBe(true);
    }
  });
  it('allows public addresses', () => {
    for (const ip of ['8.8.8.8', '142.250.183.97', '2607:f8b0:4004:800::200e']) expect(isBlockedIp(ip), ip).toBe(false);
  });
  it('only https URLs without credentials or private literals', () => {
    expect(() => assertFetchableUrl('http://lh3.googleusercontent.com/x')).toThrow(/https/);
    expect(() => assertFetchableUrl('https://user:pw@example.com/x')).toThrow();
    expect(() => assertFetchableUrl('https://127.0.0.1/x')).toThrow();
    expect(() => assertFetchableUrl('https://[::1]/x')).toThrow();
    expect(() => assertFetchableUrl('https://localhost/x')).toThrow();
    expect(() => assertFetchableUrl('https://example.com:8443/x')).toThrow();
    expect(assertFetchableUrl('https://lh3.googleusercontent.com/abc').hostname).toBe('lh3.googleusercontent.com');
  });
});
