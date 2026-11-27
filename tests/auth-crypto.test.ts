import crypto from 'node:crypto';
import { beforeAll, describe, expect, it } from 'vitest';

beforeAll(() => {
  process.env.APP_ENCRYPTION_KEY = crypto.randomBytes(32).toString('base64');
});

const load = () => import('../server/lib/crypto.js');

describe('seal / open (AES-256-GCM, AAD = sender account id)', () => {
  it('round-trips a Google App Password', async () => {
    const { seal, open } = await load();
    const id = crypto.randomUUID();
    const sealed = seal('abcdefghijklmnop', id);
    expect(sealed.enc).not.toContain('abcdefghijklmnop');
    expect(sealed.keyVersion).toBe(1);
    expect(open(sealed, id)).toBe('abcdefghijklmnop');
  });

  it('uses a fresh IV for every seal', async () => {
    const { seal } = await load();
    const a = seal('same', 'id'); const b = seal('same', 'id');
    expect(a.iv).not.toBe(b.iv);
    expect(a.enc).not.toBe(b.enc);
  });

  it('rejects a modified auth tag', async () => {
    const { seal, open } = await load();
    const sealed = seal('abcdefghijklmnop', 'acct-1');
    const tag = Buffer.from(sealed.tag, 'base64'); tag[0] ^= 0xff;
    expect(() => open({ ...sealed, tag: tag.toString('base64') }, 'acct-1')).toThrow();
  });

  it('rejects modified ciphertext', async () => {
    const { seal, open } = await load();
    const sealed = seal('abcdefghijklmnop', 'acct-1');
    const enc = Buffer.from(sealed.enc, 'base64'); enc[0] ^= 0x01;
    expect(() => open({ ...sealed, enc: enc.toString('base64') }, 'acct-1')).toThrow();
  });

  it('rejects the wrong AAD (ciphertext copied onto another account)', async () => {
    const { seal, open } = await load();
    const sealed = seal('abcdefghijklmnop', 'acct-1');
    expect(() => open(sealed, 'acct-2')).toThrow();
  });

  it('rejects a different key', async () => {
    const { seal, open } = await load();
    const sealed = seal('abcdefghijklmnop', 'acct-1');
    const old = process.env.APP_ENCRYPTION_KEY;
    process.env.APP_ENCRYPTION_KEY = crypto.randomBytes(32).toString('base64');
    try { expect(() => open(sealed, 'acct-1')).toThrow(); } finally { process.env.APP_ENCRYPTION_KEY = old; }
  });
});

describe('signToken / verifyToken', () => {
  it('round-trips a payload', async () => {
    const { signToken, verifyToken } = await load();
    const t = signToken({ e: 'a@b.com', j: 'job-1' });
    expect(verifyToken<{ e: string; j: string }>(t)).toEqual({ e: 'a@b.com', j: 'job-1' });
  });

  it('rejects a tampered body or mac', async () => {
    const { signToken, verifyToken } = await load();
    const t = signToken({ e: 'a@b.com' });
    const [body, mac] = t.split('.');
    const forged = Buffer.from(JSON.stringify({ e: 'evil@b.com' })).toString('base64url');
    expect(verifyToken(`${forged}.${mac}`)).toBeNull();
    expect(verifyToken(`${body}.${mac.slice(0, -1)}${mac.endsWith('A') ? 'B' : 'A'}`)).toBeNull();
    expect(verifyToken(body)).toBeNull();
    expect(verifyToken('')).toBeNull();
  });
});

describe('tempPassword', () => {
  it('produces three 4-character groups without ambiguous characters', async () => {
    const { tempPassword } = await load();
    for (let i = 0; i < 50; i++) {
      const p = tempPassword();
      expect(p).toMatch(/^[A-Za-z2-9]{4}-[A-Za-z2-9]{4}-[A-Za-z2-9]{4}$/);
      expect(p).not.toMatch(/[01IOlo]/);
    }
  });
});
